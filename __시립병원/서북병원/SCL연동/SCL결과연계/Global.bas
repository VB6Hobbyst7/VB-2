Attribute VB_Name = "Global"
Option Explicit

Public LoginSucceeded As Boolean
Public gConnect As Boolean

Public gCodeItem As String      '����ȭ�鿡�� ���õ� �ڵ���� �׸�
Public gSubItem As String
Public gSelectItem As String

Public gSaveFlag As Integer     '�ű� �Է� = 1, ���� = 2
Public gReturn As Integer       '�Է�Ȥ�� ���� ������ �߻�����
                                '�߻��� :  1
                                '�̹߻� : -1
                                
Public gPID As String
Public gReceNo As String
Public gCode As String
Public gName As String
Public gSeqNo As String

Public gUID As String
Public gUName As String

Public gDateTime As String

Public gReason As String


Public gWEgb As String  '�����˻�=1, ����˻�=2
Public gSection As String

Public iToggle As Integer  '�˻� ��� �Է� �� toggle

'Command Botton
'-------------------------------------------------------
 ' ���ڷ��Է� : 0
 ' �ڷ� ����  : 1
 ' �ڷ� ����  : 2
 ' �ڷ� ����  : 3
 ' �� ��      : 4
 ' �ڷ� ���  : 5
 ' ������     : 6
 ' Hold       : 7
'--------------------------------------------------------
Public Const cmdAddNew = 0
Public Const cmdEdit = 1
Public Const cmdDelete = 2
Public Const cmdSave = 3
Public Const cmdCancel = 4
Public Const cmdPrint = 5
Public Const cmdExit = 6
Public Const cmdHold = 7



Public Type Hospital
    HID         As String
    HName       As String
    HNumber     As String
    Address     As String
    Catagory    As String
    Business    As String
    Head        As String
    Phone       As String
    License1st  As String
    Doctor1st   As String
    LabLicense  As String
    LabDoctor   As String
End Type
Public gHosInfo As Hospital

Public Type Patient
    PID         As String
    PName       As String
    Address     As String   'Patient�� Address �� Street�� ��ģ��
    Jumin1      As String
    Jumin2      As String
    Phone       As String
    Sex         As String
    Age         As String
    InsType     As String
    ABOType     As String
    RHType      As String
End Type
Public gPatInfo As Patient

Public Type Dept           ' �μ� ��ü(C,T,E)
    Code        As String
    EName       As String
    KName       As String
    Alias       As String
    Gubun       As String
    UseFlag     As String
    Remark      As String
End Type
Public gDept As Dept
    

Public Type CDept
    Code        As String
    EName       As String
    KName       As String
    Alias       As String
End Type
Public gCDept As CDept      '����� ���� Dept�� Gubun �� 'C' �ΰ�

Public Type TxDept
    Code        As String
    EName       As String
    KName       As String
    Alias       As String
End Type
Public gTxDept As TxDept      '���������μ� ���� Dept�� Gubun �� 'T' �ΰ�

Public Type EDept
    Code        As String
    EName       As String
    KName       As String
    Alias       As String
End Type
Public gEDept As EDept      '��Ÿ �μ� ���� Dept�� Gubun �� 'E' �ΰ�


Public Type User
    ID         As String
    GradeCode   As String
    Name       As String
    Passwd     As String
    Position   As String
    Jumin1      As String
    Jumin2      As String
    DeptCode    As String
    ClassCode   As String
    WardCode    As String
    EmgCall     As String
    Remark     As String
    BankUse     As String
    CodeUse     As String
    MicUse      As String
    ResultUse   As String
End Type
Public gUser As User        '���α׷� ����� ����

Public Type Slip
    Code        As String
    Name        As String
    Alias       As String
End Type
Public gSlip As Slip        'SlipClass ����

Public Type LabMaster
    Code        As String
    Name        As String
    RoomNo      As String
    Phone       As String
End Type
Public gLab As LabMaster          '�ӻ󺴸��� ��Ʈ ����


Public Type CaseMaster
    Code        As String
    Name        As String
    Remark      As String
End Type
Public gCase As CaseMaster      ' ��⸶����

Public Type UnitMaster
    Code        As String
    Name        As String
End Type
Public gUnit As UnitMaster      ' ����������

Public Type Exam
    Code        As String
    EName       As String
    Aliase      As String
    KName       As String
    SpecimenVol As String
    SpecimenCode    As String
    SpecimenName    As String
    SpecimenID  As String
    PickUID     As String
    PickDate    As String
    Method      As String
    SlipCode    As String
    LabCode     As String
    EquipCode   As String
    CaseCode    As String
    WorkDay     As String
    ReqDay      As String
    Unit        As String
    ResHigh     As String
    ResLow      As String
    PanicValueGubun As String
    PanicHigh   As String
    PanicLow    As String
    DeltaValueGubun As String
    DeltaHigh   As String
    DeltaLow    As String
    EtcRefFlag  As String
    EtcDeltaFlag    As String
End Type
Public gExam As Exam        '�˻� �׸� ����

Public Type Trust
    Code        As String
    Flag        As String
    Name        As String
    BizCode     As String
    Address     As String
    HeadName    As String
    Charge      As String
    Phone1      As String
    Phone2      As String
    Phone3      As String
    Email       As String
    Remark      As String
End Type
Public gTrust As Trust      '�ŷ�ó ����

Public Type Equip
    Code        As String
    Name        As String
    Maker       As String
    Admin       As String
    Phone1        As String
    Phone2        As String
    Phone3        As String
    BuyDate     As String
    UseFlag     As String
    Remark      As String
End Type
Public gEquip As Equip      ' ���

Public Type Ward
    Code        As String
    Room        As String
    Name        As String
    BedCheck    As String
    Nurse       As String
    UseFlag     As String
    Remark      As String
End Type
Public gWard As Ward        '����


    
    


