VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisMasterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ� : clsLisMasterForm.cls
'   �ۼ��� :
'   ��  �� :
'   �ۼ��� :
'   ��  �� :
'       1. 5.0.4: �̻��(2005-01-04)
'          - Ʈ���信 WBC �˻��׸���, NRBC �˻��׸��� �޴��߰�
'-----------------------------------------------------------------------------'

Option Explicit

Private WithEvents fL340    As Form    'formfrm340ShowIndex
Attribute fL340.VB_VarHelpID = -1
Private WithEvents fL340_1  As Form    'formfrm340ShowIndex2
Attribute fL340_1.VB_VarHelpID = -1
Private WithEvents fL341    As Form    'frm341Common1
Attribute fL341.VB_VarHelpID = -1
Private WithEvents fL342    As Form    'frm342Common2
Attribute fL342.VB_VarHelpID = -1
Private WithEvents fL343    As Form    'frm343Template
Attribute fL343.VB_VarHelpID = -1

Private WithEvents fL351    As Form    'frm351ItemMaster
Attribute fL351.VB_VarHelpID = -1
Private WithEvents fL352    As Form    'frm352Specimen
Attribute fL352.VB_VarHelpID = -1
Private WithEvents fL353    As Form    'frm353Reference
Attribute fL353.VB_VarHelpID = -1
Private WithEvents fL362    As Form    'frm362WSMaster
Attribute fL362.VB_VarHelpID = -1
Private WithEvents fL363    As Form    'frm363MicWsKind
Attribute fL363.VB_VarHelpID = -1
Private WithEvents fL364    As Form    'frm364AppAnti
Attribute fL364.VB_VarHelpID = -1
Private WithEvents fL365    As Form    'frm365SpcGroup
Attribute fL365.VB_VarHelpID = -1
Private WithEvents fL366    As Form    'frm366EDefine
Attribute fL366.VB_VarHelpID = -1
Private WithEvents fL370    As Form    'frm370CumCdSet
Attribute fL370.VB_VarHelpID = -1
Private WithEvents fL381    As Form    'frm381EquipMaster
Attribute fL381.VB_VarHelpID = -1
Private WithEvents fL382    As Form    'frm382BasePrint
Attribute fL382.VB_VarHelpID = -1
Private WithEvents fL601    As Form    'frm601MachHistory
Attribute fL601.VB_VarHelpID = -1

Private WithEvents fL383    As Form    'frm383
Attribute fL383.VB_VarHelpID = -1
Private WithEvents fL384    As Form    'frm384
Attribute fL384.VB_VarHelpID = -1
Private WithEvents fL385    As Form    'frm385
Attribute fL385.VB_VarHelpID = -1
Private WithEvents fL386    As Form    'frm386
Attribute fL386.VB_VarHelpID = -1

Private WithEvents fL388    As Form    'frm388InterfaceEquip
Attribute fL388.VB_VarHelpID = -1

Private WithEvents fL602    As Form    'frm602OnlineHelp
Attribute fL602.VB_VarHelpID = -1
Private CurNode             As Node

'Public Property Let IsDeveloper(ByVal vData As Boolean)
'    gIsDeveloper = vData
'End Property
'
'Public Property Let EmpId(ByVal vData As String)
'    gEmpId = vData
'End Property

Private Function GetUseFg(ByVal qTmp As String) As String
    Dim strPath As String
    
    If Dir(INIPath) = "" Then Exit Function
    
    GetUseFg = medGetINI("LIS_MASTER", qTmp, INIPath)
    GetUseFg = medGetP(GetUseFg, 2, LINE_DIV)
End Function

Public Sub MasterTreeviewLoad(tvwMenu As Object)
    Dim strUseFg    As String
    
    With tvwMenu
        .Nodes.Clear
        If ObjSysInfo.EmpId = "9999" Then
            If GetUseFg("L0") = "0" Then _
                Call .Nodes.Add(, , "L0", "LIS �ڵ� ����", 1)
            If GetUseFg("LC2_Index") = "0" Then _
                Call .Nodes.Add("L0", tvwChild, LC2_Index, "�ڵ�-1 Index", 2)
            If GetUseFg("LC3_Index") = "0" Then _
                Call .Nodes.Add("L0", tvwChild, LC3_Index, "�ڵ�-2 Index", 2)
            If GetUseFg("LC4_Index") = "0" Then _
                Call .Nodes.Add("L0", tvwChild, LC4_Index, "Template Index", 2)
            If GetUseFg("LC5_Index") = "0" Then _
                Call .Nodes.Add("L0", tvwChild, LC5_Index, "������Ʈ ����", 2)
            If GetUseFg("M006") = "0" Then _
                Call .Nodes.Add("L0", tvwChild, "M006", "�������̽� ��� ����", 2)
            tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
            If GetUseFg("L1") = "0" Then _
                Call .Nodes.Add("L0", tvwNext, "L1", "�˻��׸� ������", 1)
        Else
            If GetUseFg("L1") = "0" Then _
                Call .Nodes.Add(, , "L1", "�˻��׸� ������", 1)
        End If
        
        If GetUseFg("M201") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, "M201", "�˻��׸� �ڵ�", 2)
        If GetUseFg("M202") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, "M202", "������ü ����", 2)
        If GetUseFg("M203") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, "M203", "����ġ ����", 2)
        If GetUseFg("LC2_Detail") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, LC2_Detail, "���׸� ����", 2)
        If GetUseFg("LC3_WorkArea") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, LC3_WorkArea, "Work Area", 2)
        If GetUseFg("LC3_Specimen") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, LC3_Specimen, "��ü �ڵ�", 2)
        If GetUseFg("LC2_MultiSpc") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, LC2_MultiSpc, "������ü", 2)
        If GetUseFg("LC3_StoreCd") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, LC3_StoreCd, "�������� �ڵ�", 2)
        If GetUseFg("M305") = "0" Then _
            Call .Nodes.Add("L1", tvwChild, "M305", "�����ڷ� ��ȸ", 2)
            
        ' �������� �߰� By M.G.Choi 2004.11.01 =================================
        Call .Nodes.Add("L1", tvwChild, LC3_INFECTION, "�������� �Ƿڰ�ü", 2)
        Call .Nodes.Add("L1", tvwChild, LC3_INFECTIONTEST, "�������� �������", 2)
        
        Call .Nodes.Add("L1", tvwChild, LC3_RESULTREADTEST, "�����ǵ� �˻��׸�", 2)
        '=======================================================================
        
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible

        If GetUseFg("L2") = "0" Then _
            Call .Nodes.Add("L1", tvwNext, "L2", "ó�� ���� �ڵ�", 1)
        If GetUseFg("LC3_HighItem") = "0" Then _
            Call .Nodes.Add("L2", tvwChild, LC3_HighItem, "�ٺ� ó��", 2)
        If GetUseFg("LC3_ByPass") = "0" Then _
            Call .Nodes.Add("L2", tvwChild, LC3_ByPass, "ByPass ó��", 2)
        If GetUseFg("LC2_Panel") = "0" Then _
            Call .Nodes.Add("L2", tvwChild, LC2_Panel, "Panel ó��", 2)
        If GetUseFg("LC2_AttrItem") = "0" Then _
            Call .Nodes.Add("L2", tvwChild, LC2_AttrItem, "�˻��׸� �Ӽ�", 2)
        If GetUseFg("LC3_ICUTestCd") = "0" Then _
            Call .Nodes.Add("L2", tvwChild, LC3_ICUTestCd, "ICU �˻��׸� ����", 2)
        If GetUseFg("LC3_POCTestCd") = "0" Then _
            Call .Nodes.Add("L2", tvwChild, LC3_POCTestCd, "���� �˻��׸� ����", 2)
        If GetUseFg("LC3_POCTimes") = "0" Then _
            Call .Nodes.Add("L2", tvwChild, LC3_POCTimes, "���� �˻�ð��� ����", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L3") = "0" Then _
            Call .Nodes.Add("L2", tvwNext, "L3", "ä��/���� ���� �ڵ�", 1)
        If GetUseFg("LC3_ColTeam") = "0" Then _
            Call .Nodes.Add("L3", tvwChild, LC3_ColTeam, "ä���� ���", 2)
        If GetUseFg("LC3_RoundTime") = "0" Then _
            Call .Nodes.Add("L3", tvwChild, LC3_RoundTime, "Rounding ä�� �ð���", 2)
        If GetUseFg("LC3_WorkTime") = "0" Then _
            Call .Nodes.Add("L3", tvwChild, LC3_WorkTime, "�����ð� �ۼ�", 2)
        If GetUseFg("LC3_OutLab") = "0" Then _
            Call .Nodes.Add("L3", tvwChild, LC3_OutLab, "�ܺ��Ƿ�ó", 2)
        If GetUseFg("LC3_RefLab") = "0" Then _
            Call .Nodes.Add("L3", tvwChild, LC3_RefLab, "Referral Lab", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L4") = "0" Then _
            Call .Nodes.Add("L3", tvwNext, "L4", "��� ��� ���� �ڵ�", 1)
        If GetUseFg("M207") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, "M207", "Work sheet ������", 2)
        If GetUseFg("LC2_RelTest") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC2_RelTest, "���� �˻��׸� ����", 2)
        If GetUseFg("LC2_Calculation") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC2_Calculation, "������ ����", 2)

        If GetUseFg("LC2_WAResult") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC2_WAResult, "WorkArea�� ��� �ڵ�", 2)
        If GetUseFg("LC2_ItemResult") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC2_ItemResult, "�˻��׸� ��� �ڵ�", 2)
        If GetUseFg("LC3_ElectronicSign") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC3_ElectronicSign, "���ڼ��� ����", 2)
        If GetUseFg("LC3_TUBERCLE") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC3_TUBERCLE, "���ٱհ˻� �����ڵ�", 2)

        If GetUseFg("LC2_TempletTest") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC2_TempletTest, "Ư���˻� ���ø� �Ϲݰ˻��ڵ�", 2)
        If GetUseFg("LC2_TempletText1") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC2_TempletText1, "Ư���˻� ���ø� �ؽ�Ʈ1", 2)
        If GetUseFg("LC2_TempletText2") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC2_TempletText2, "Ư���˻� ���ø� �ؽ�Ʈ2", 2)
        If GetUseFg("LC2_TempletText3") = "0" Then _
            Call .Nodes.Add("L4", tvwChild, LC2_TempletText3, "Ư���˻� ���ø� �ؽ�Ʈ3", 2)
        
        '** ������ Ȯ�� �˻��׸� ��� By M.G.Choi 2006.09.01
        Call .Nodes.Add("L4", tvwChild, LC2_DoctTest, "������ Ȯ�� �˻��׸� ����", 2)
        
        ' �������� �߰� By M.G.Choi 2004.08.23 =================================
        Call .Nodes.Add("L4", tvwChild, LC4_Infection, "�������� �˻���", 2)
        '=======================================================================
        ' �������� �߰� By M.G.Choi 2007.07.02 =================================
        Call .Nodes.Add("L4", tvwChild, LC2_Warning, "Warning Message", 2)
        '=======================================================================
        
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L5") = "0" Then _
            Call .Nodes.Add("L4", tvwNext, "L5", "�̻���/Ư���˻� ���� �ڵ�", 1)
        If GetUseFg("M301") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, "M301", "��ü�� ���", 2)
        If GetUseFg("LC3_Media") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, LC3_Media, "�����ڵ�", 2)
        If GetUseFg("M302") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, "M302", "�̻��� ��ġ ��� ���", 2)
        If GetUseFg("LC3_Microbe") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, LC3_Microbe, "���ڵ�", 2)
        If GetUseFg("LC3_Species") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, LC3_Species, "�����ڵ�", 2)
        If GetUseFg("LC3_AntiBiotic") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, LC3_AntiBiotic, "�׻����ڵ�", 2)
        If GetUseFg("LC3_Volume") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, LC3_Volume, "�����ڵ�", 2)
        If GetUseFg("M303") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, "M303", "������ �����׻��� ����", 2)
        If GetUseFg("LC3_MWSKinds") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, LC3_MWSKinds, "�̻��� ��ũ��Ʈ ����", 2)
        If GetUseFg("M304") = "0" Then _
            Call .Nodes.Add("L5", tvwChild, "M304", "Ư�� �˻� ȭ�� ����", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L6") = "0" Then _
            Call .Nodes.Add("L5", tvwNext, "L6", "�����ȸ/��� ���� �ڵ�", 1)
        If GetUseFg("LC3_StaticItem") = "0" Then _
            Call .Nodes.Add("L6", tvwChild, LC3_StaticItem, "���� ��� �׸� ����", 2)
        If GetUseFg("LC3_StaticGroup") = "0" Then _
            Call .Nodes.Add("L6", tvwChild, LC3_StaticGroup, "���� ��� Group", 2)
        If GetUseFg("M401") = "0" Then _
            Call .Nodes.Add("L6", tvwChild, "M401", "�����ڵ� ���", 2)
        If GetUseFg("LC3_StartDate") = "0" Then _
            Call .Nodes.Add("L6", tvwChild, LC3_StartDate, "���� ����˻��Ⱓ ����", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L7") = "0" Then _
            Call .Nodes.Add("L6", tvwNext, "L7", "Template �ڵ� ����", 1)
        If GetUseFg("LC4_FootNote") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_FootNote, "Foot Note", 2)
        If GetUseFg("LC4_Remark") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_Remark, "��ü Remark", 2)
        If GetUseFg("LC4_WarnInfect") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_WarnInfect, "Warning/Infection", 2)
        If GetUseFg("LC4_TextResult") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_TextResult, "Text ���", 2)
        If GetUseFg("LC4_ClinicalNotice") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_ClinicalNotice, "Clinical Notice", 2)
        If GetUseFg("LC4_CancelReason") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_CancelReason, "������� & ��ä�� ����", 2)
        If GetUseFg("LC4_ModifyReason") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_ModifyReason, "������� ����", 2)
        If GetUseFg("LC4_QCRejReason") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_QCRejReason, "QC Reject ����", 2)
        If GetUseFg("LC4_Calibration") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_Calibration, "QC Calibration �Ұ�", 2)
        
        If GetUseFg("LC4_AccessComment") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_AccessComment, "�ܷ������� Comment ���ø�", 2)
        If GetUseFg("LC4_TestItemComment") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_TestItemComment, "�˻��׸� Comment", 2)
        
        If GetUseFg("LC4_UncolReason") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_UncolReason, "��ä�� ����", 2)
        If GetUseFg("LC4_DCReason") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_DCReason, "ó����� ����", 2)
        If GetUseFg("LC4_Peripheral") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_Peripheral, "Peripheral ���ø�", 2)
        If GetUseFg("LC4_BMComment") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_BMComment, "Bone Marrow Comment", 2)
        If GetUseFg("LC4_Hematologic") = "0" Then _
            Call .Nodes.Add("L7", tvwChild, LC4_Hematologic, "Hematologic ���ø�", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L8") = "0" Then _
            Call .Nodes.Add("L7", tvwNext, "L8", "Quality Control", 1)
        If GetUseFg("LC3_PartCd") = "0" Then _
            Call .Nodes.Add("L8", tvwChild, LC3_PartCd, "�ܺ��������� �а��ڵ�", 2)
        If GetUseFg("LC3_HospCd") = "0" Then _
            Call .Nodes.Add("L8", tvwChild, LC3_HospCd, "�Ƿ��� �����ڵ�", 2)
        If GetUseFg("LC3_Method") = "0" Then _
            Call .Nodes.Add("L8", tvwChild, LC3_Method, "�˻��� �ڵ�", 2)
        If GetUseFg("LC2_MicQcRst") = "0" Then _
            Call .Nodes.Add("L8", tvwChild, LC2_MicQcRst, "�̻���QC ����ڵ�", 2)
        If GetUseFg("LC3_BldQcRst") = "0" Then _
            Call .Nodes.Add("L8", tvwChild, LC3_BldQcRst, "��������QC ����ڵ�", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L9") = "0" Then _
            Call .Nodes.Add("L8", tvwNext, "L9", "��  Ÿ", 1)
        If GetUseFg("LC3_Buildings") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_Buildings, "�ǹ��ڵ� ���", 2)
        If GetUseFg("LC3_Section") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_Section, "Section �ڵ�", 2)
        If GetUseFg("M206") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, "M206", "��� ������", 2)
        If GetUseFg("LC3_Vander") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_Vander, "��ü ������", 2)
        If GetUseFg("LC3_PtDiv") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_PtDiv, "ȯ�ڱ��� �ڵ�", 2)
        If GetUseFg("LC3_ReportTesctCd") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_ReportTesctCd, "ȸ���뷹��Ʈ �˻��׸� ����", 2)
        If GetUseFg("LC3_WBCDiffCode") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_WBCDiffCode, "WBC Diff �˻��ڵ�", 2)
        If GetUseFg("LC3_DiffKeyMap") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_DiffKeyMap, "WBC Diff KeyMap �˻��ڵ�", 2)
        If GetUseFg("LC3_WBCCode") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_WBCCode, "WBC �˻��ڵ� ����", 2)
        If GetUseFg("LC3_NRBCCode") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_NRBCCode, "NRBC �˻��ڵ� ����", 2)
        If GetUseFg("LC3_DiffCount") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_DiffCount, "Bone Marrow Diff Count�ڵ�", 2)
        
        If GetUseFg("LC3_BatchColDept") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_BatchColDept, "�ܷ��ϰ�ä�� ����μ�", 2)
'        -- �߰� By MGChoi 2002.08.16
        If GetUseFg("LC3_MachHistory") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC3_MachHistory, "��������ڵ�", 2)
'        If GetUseFg("M601") = "0" Then _
'        Call .Nodes.Add("L9", tvwChild, "M601", "�����°���", 2)
        If GetUseFg("LC2_OGYITEM") = "0" Then _
            Call .Nodes.Add("L9", tvwChild, LC2_OGYITEM, "����ΰ� �˻��׸��ڵ�", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L10") = "0" Then _
            Call .Nodes.Add("L9", tvwNext, "L10", "ȯ�� ����", 1)
        If GetUseFg("LC3_FileServer") = "0" Then _
            Call .Nodes.Add("L10", tvwChild, LC3_FileServer, "File Server ��ġ", 2)
        If GetUseFg("LC3_PrinterId") = "0" Then _
            Call .Nodes.Add("L10", tvwChild, LC3_PrinterId, "������ID ���", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible
        
        If GetUseFg("L11") = "0" Then _
            Call .Nodes.Add("L10", tvwNext, "L11", "�ܺμ�Ź�˻�", 1)
        If GetUseFg("OT1") = "0" Then _
            Call .Nodes.Add("L11", tvwChild, "OT1", "�ܺμ�Ź����_Match", 2)
        If GetUseFg("OT2") = "0" Then _
            Call .Nodes.Add("L11", tvwChild, "OT2", "�ܺμ�Ź�˻缳��", 2)
        If GetUseFg("OT3") = "0" Then _
            Call .Nodes.Add("L11", tvwChild, "OT3", "�ܺμ�Ź�˻���ȸ", 2)
        If GetUseFg("OT4") = "0" Then _
            Call .Nodes.Add("L11", tvwChild, "OT4", "�ܺμ�Ź�ǷڰǼ�", 2)
        tvwMenu.Nodes(tvwMenu.Nodes.Count).EnsureVisible

      .BorderStyle = vbFixedSingle
    
    End With
   
    tvwMenu.Nodes(1).Selected = True
    Set CurNode = tvwMenu.Nodes(1)

End Sub

Public Sub MasterTreeviewNodeClick(ByVal Rkey As String, ByVal RName As String, Optional ByVal picForm As Object)

On Error GoTo MasterTreeviewNodeClick_error

    gParentWhnd = picForm.hWnd
    
    Select Case Rkey
        Case LC2_Index:
            If fL340 Is Nothing Then Set fL340 = frm340ShowIndex
            Call SetParent(fL340.hWnd, picForm.hWnd)
            fL340.Caption = RName
            fL340.WindowState = 2
            DoEvents
            fL340.Show
            fL340.optSelect(0) = True
            fL340.ZOrder
            DoEvents
        Case LC3_Index:
            If fL340 Is Nothing Then Set fL340 = frm340ShowIndex
            Call SetParent(fL340.hWnd, picForm.hWnd)
            fL340.Caption = RName
            fL340.WindowState = 2
            DoEvents
            fL340.Show
            fL340.optSelect(1) = True
            fL340.ZOrder
            DoEvents
        Case LC4_Index:
            If fL340 Is Nothing Then Set fL340 = frm340ShowIndex
            Call SetParent(fL340.hWnd, picForm.hWnd)
            fL340.Caption = RName
            fL340.WindowState = 2
            DoEvents
            fL340.Show
            fL340.optSelect(2) = True
            fL340.ZOrder
            DoEvents
        Case LC5_Index:
            If fL340_1 Is Nothing Then Set fL340_1 = frm340ShowIndex2
            Call SetParent(fL340_1.hWnd, picForm.hWnd)
            fL340_1.Caption = RName
            fL340_1.WindowState = 2
            DoEvents
            fL340_1.Show
            fL340_1.ZOrder
            DoEvents
            
        Case "M201":
            If fL351 Is Nothing Then Set fL351 = frm351ItemMaster
            Call SetParent(fL351.hWnd, picForm.hWnd)
            fL351.Caption = RName
            fL351.WindowState = 2
            DoEvents
            fL351.Show
            fL351.ZOrder
            DoEvents
        Case "M202":
            If fL352 Is Nothing Then Set fL352 = frm352Specimen
            Call SetParent(fL352.hWnd, picForm.hWnd)
            fL352.Caption = RName
            fL352.WindowState = 2
            DoEvents
            fL352.Show
            fL352.ZOrder
            DoEvents
        Case "M203":
            If fL353 Is Nothing Then Set fL353 = frm353Reference
            Call SetParent(fL353.hWnd, picForm.hWnd)
            fL353.Caption = RName
            fL353.WindowState = 2
            DoEvents
            fL353.Show
            fL353.ZOrder
            DoEvents
        Case "M207":
            If fL362 Is Nothing Then Set fL362 = frm362WSMaster
            Call SetParent(fL362.hWnd, picForm.hWnd)
            fL362.Caption = RName
            fL362.WindowState = 2
            DoEvents
            fL362.Show
            fL362.ZOrder
            DoEvents
        Case "M301":
            If fL365 Is Nothing Then Set fL365 = frm365SpcGroup
            Call SetParent(fL365.hWnd, picForm.hWnd)
            fL365.Caption = RName
            fL365.WindowState = 2
            DoEvents
            fL365.Show
            fL365.ZOrder
            DoEvents
        Case "M302":
            If fL363 Is Nothing Then Set fL363 = frm363MicWsKind
            Call SetParent(fL363.hWnd, picForm.hWnd)
            fL363.Caption = RName
            fL363.WindowState = 2
            DoEvents
            fL363.Show
            fL363.ZOrder
            DoEvents
        Case "M305"
            If fL382 Is Nothing Then Set fL382 = frm382BasePrint
            Call SetParent(fL382.hWnd, picForm.hWnd)
            fL382.Caption = RName
            fL382.WindowState = 2
            fL382.Show
            fL382.ZOrder
            DoEvents
        Case "M303":
            If fL364 Is Nothing Then Set fL364 = frm364AppAnti
            Call SetParent(fL364.hWnd, picForm.hWnd)
            fL364.Caption = RName
            fL364.WindowState = 2
            DoEvents
            fL364.Show
            fL364.ZOrder
            DoEvents
        Case "M304":
            If fL366 Is Nothing Then Set fL366 = frm366EDefine
            Call SetParent(fL366.hWnd, picForm.hWnd)
            fL366.Caption = RName
            fL366.WindowState = 2
            DoEvents
            fL366.Show
            fL366.ZOrder
            DoEvents
        Case "M401":
            If fL370 Is Nothing Then Set fL370 = frm370CumCdSet
            Call SetParent(fL370.hWnd, picForm.hWnd)
            fL370.Caption = RName
            fL370.WindowState = 2
            fL370.IsManager = True
            fL370.DeptCd = ""
            DoEvents
            fL370.Show
            fL370.ZOrder
            DoEvents
        Case "M206":
            If fL381 Is Nothing Then Set fL381 = frm381EquipMaster
            Call SetParent(fL381.hWnd, picForm.hWnd)
            fL381.Caption = RName
            fL381.WindowState = 2
            fL381.Show
            DoEvents
            fL381.ZOrder
            DoEvents
'        Case "M601":
'            If fL601 Is Nothing Then Set fL601 = frm601MachHistory
'            Call SetParent(fL601.hWnd, picForm.hWnd)
'            fL601.Caption = RName
'            fL601.WindowState = 2
'            fL601.Show
'            DoEvents
'            fL601.ZOrder
'            DoEvents
        Case LC2_Panel, LC2_MultiSpc, LC2_Detail, LC2_QcControl, LC2_KeyMap, LC2_AttrItem, _
             LC2_MediaBio, LC2_MicroAnti, LC2_Machine, LC2_ItemResult, LC2_WAResult, LC2_RelTest, _
             LC2_MicQcRst, LC2_Calculation, LC2_OGYITEM, LC2_TempletTest, LC2_TempletText1, _
             LC2_TempletText2, LC2_TempletText3, LC2_DoctTest, LC2_Warning
            
            If fL341 Is Nothing Then Set fL341 = frm341Common1
            Call SetParent(fL341.hWnd, picForm.hWnd)
            fL341.Caption = RName
            fL341.Rkey = Rkey
            DoEvents
            fL341.RName = RName 'T_COM002 & ";" & RName
            fL341.WindowState = 2
            DoEvents
            fL341.Show
            fL341.ZOrder
                 
        Case LC3_DrGrade, LC3_BedGrade, LC3_BedStatus, LC3_DeptDiv, _
             LC3_HighItem, LC3_PocItem, LC3_ByPass, LC3_RoundTime, LC3_ColTeam, _
             LC3_OutLab, LC3_RefLab, LC3_Vander, LC3_Section, LC3_WorkArea, LC3_Section, _
             LC3_Specimen, LC3_VerifyFg, LC3_Media, LC3_Microbe, LC3_Species, _
             LC3_AntiBiotic, LC3_BioChemical, LC3_Volume, LC3_Infect, LC3_Buildings, _
             LC3_StoreCd, LC3_FileServer, LC3_StaticItem, LC3_StaticGroup, LC3_PrinterId, _
             LC3_StartDate, LC3_PtDiv, LC3_PartCd, LC3_HospCd, LC3_Method, LC3_BldQcRst, _
             LC3_ICUTestCd, LC3_POCTestCd, LC3_POCTimes, LC3_ElectronicSign, LC3_ReportTesctCd, _
             LC3_WBCDiffCode, LC3_DiffKeyMap, LC3_BatchColDept, LC3_DiffCount, "C252", LC3_TUBERCLE, LC3_MWSKinds, LC3_WorkTime, _
             LC3_INFECTION, LC3_INFECTIONTEST, LC3_RESULTREADTEST, LC3_WBCCode, LC3_NRBCCode
             
            If fL342 Is Nothing Then Set fL342 = frm342Common2
            Call SetParent(fL342.hWnd, picForm.hWnd)
            
            fL342.Rkey = Rkey
            DoEvents
            fL342.RName = RName & ":" & T_COM003
            fL342.Caption = RName
            fL342.WindowState = 2
            DoEvents
            fL342.Show
            fL342.ZOrder
                 
'        Case LC4_ScrLock, LC4_PrgOnOff, LC4_FnctOnOff, LC4_InfectCond, _
'             LC4_BarFormat, LC4_BarTime, LC4_WSPrtTime, LC4_Hospital, LC4_CumulTime, _
'             LC4_LabelTime, LC4_TempUnit, LC4_DateFormat, LC4_TimeFormat
            'frm356CCommon2.RKey = Node.Key
            'frm356CCommon2.RName = Node.Text
            'frm356CCommon2.Show
            'frm356CCommon2.zorder
        'LC5_StainResult
        Case LC4_TestItemComment
        
            If fL602 Is Nothing Then Set fL602 = frm602OnlineHelp
            Call SetParent(fL602.hWnd, picForm.hWnd)
            fL602.Caption = RName
            fL602.WindowState = 2
            fL602.Show
            DoEvents
            fL602.ZOrder
            DoEvents
        
        Case LC4_Morphology, LC4_UncolReason, LC4_Remark, _
             LC4_FootNote, LC4_WarnInfect, LC4_TextResult, LC4_SPTextResult, LC4_DCReason, _
             LC4_CancelReason, LC4_ModifyReason, LC4_QCRejReason, LC4_TempReason, LC4_ClinicalNotice, _
             LC4_Peripheral, LC4_BMComment, LC4_Hematologic, LC4_TempletCommet, LC4_TempletDignosis, _
             LC4_Calibration, LC4_AccessComment, LC4_Infection
            
            If fL343 Is Nothing Then Set fL343 = frm343Template
            Call SetParent(fL343.hWnd, picForm.hWnd)
            fL343.Rkey = Rkey
            DoEvents
            fL343.RName = RName
            fL343.Caption = RName
            fL343.WindowState = 2
            DoEvents
            fL343.Show
            fL343.ZOrder
        Case "OT1"
            If fL383 Is Nothing Then Set fL383 = frm383OTMatch
            Call SetParent(fL383.hWnd, picForm.hWnd)
            fL383.Caption = RName
            fL383.WindowState = 2
            DoEvents
            fL383.Show
            fL383.ZOrder
        Case "OT2"
            If fL384 Is Nothing Then Set fL384 = New frm384OtTestcd
            Call SetParent(fL384.hWnd, picForm.hWnd)
            fL384.Caption = RName
            fL384.WindowState = 2
            DoEvents
            fL384.Show
            fL384.ZOrder
        Case "OT3"
            If fL385 Is Nothing Then Set fL385 = frm385OTQuery
            Call SetParent(fL385.hWnd, picForm.hWnd)
            fL385.Caption = RName
            fL385.WindowState = 2
            DoEvents
            fL385.Show
            fL385.ZOrder
        Case "OT4"
            If fL386 Is Nothing Then Set fL386 = frm386OtPtCnt
            Call SetParent(fL386.hWnd, picForm.hWnd)
            fL386.Caption = RName
            fL386.WindowState = 2
            DoEvents
            fL386.Show
            fL386.ZOrder
        Case "M006" '�������̽� ��� ����ȭ��
            If fL388 Is Nothing Then Set fL388 = frm388InterfaceEquip
            Call SetParent(fL388.hWnd, picForm.hWnd)
            fL388.Caption = RName
            fL388.WindowState = 2
            DoEvents
            fL388.Show
        End Select
    
    Exit Sub
    
MasterTreeviewNodeClick_error:
    MsgBox Err.Description, vbCritical, "���α׷� �ε�"
    
End Sub

Public Sub MasterClose()

End Sub

Private Sub Class_Terminate()
    
    If Not (fL340 Is Nothing) Then Unload fL340: Set fL340 = Nothing
    If Not (fL340_1 Is Nothing) Then Unload fL340_1: Set fL340_1 = Nothing
    If Not (fL341 Is Nothing) Then Unload fL341: Set fL341 = Nothing
    If Not (fL342 Is Nothing) Then Unload fL342: Set fL342 = Nothing
    If Not (fL343 Is Nothing) Then Unload fL343: Set fL343 = Nothing
    
    If Not (fL351 Is Nothing) Then Unload fL351: Set fL351 = Nothing
    If Not (fL352 Is Nothing) Then Unload fL352: Set fL352 = Nothing
    If Not (fL353 Is Nothing) Then Unload fL353: Set fL353 = Nothing
    If Not (fL362 Is Nothing) Then Unload fL362: Set fL362 = Nothing
    If Not (fL363 Is Nothing) Then Unload fL363: Set fL363 = Nothing
    If Not (fL364 Is Nothing) Then Unload fL364: Set fL364 = Nothing
    If Not (fL365 Is Nothing) Then Unload fL365: Set fL365 = Nothing
    If Not (fL366 Is Nothing) Then Unload fL366: Set fL366 = Nothing
    If Not (fL370 Is Nothing) Then Unload fL370: Set fL370 = Nothing
    If Not (fL381 Is Nothing) Then Unload fL381: Set fL381 = Nothing
    
    If Not (fL382 Is Nothing) Then Unload fL382: Set fL382 = Nothing
    If Not (fL383 Is Nothing) Then Unload fL383: Set fL383 = Nothing
    If Not (fL384 Is Nothing) Then Unload fL384: Set fL384 = Nothing
    If Not (fL385 Is Nothing) Then Unload fL385: Set fL385 = Nothing
    If Not (fL386 Is Nothing) Then Unload fL386: Set fL386 = Nothing
    
    If Not (fL388 Is Nothing) Then Unload fL388: Set fL386 = Nothing
    If Not (fL602 Is Nothing) Then Unload fL602: Set fL602 = Nothing
    
    Set lstItemList = Nothing
    
End Sub