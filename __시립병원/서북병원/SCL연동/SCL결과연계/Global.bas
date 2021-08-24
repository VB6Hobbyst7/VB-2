Attribute VB_Name = "Global"
Option Explicit

Public LoginSucceeded As Boolean
Public gConnect As Boolean

Public gCodeItem As String      '메인화면에서 선택된 코드관리 항목
Public gSubItem As String
Public gSelectItem As String

Public gSaveFlag As Integer     '신규 입력 = 1, 수정 = 2
Public gReturn As Integer       '입력혹은 수정 사항의 발생사항
                                '발생시 :  1
                                '미발생 : -1
                                
Public gPID As String
Public gReceNo As String
Public gCode As String
Public gName As String
Public gSeqNo As String

Public gUID As String
Public gUName As String

Public gDateTime As String

Public gReason As String


Public gWEgb As String  '수질검사=1, 전기검사=2
Public gSection As String

Public iToggle As Integer  '검사 결과 입력 시 toggle

'Command Botton
'-------------------------------------------------------
 ' 새자료입력 : 0
 ' 자료 수정  : 1
 ' 자료 삭제  : 2
 ' 자료 저장  : 3
 ' 취 소      : 4
 ' 자료 출력  : 5
 ' 나가기     : 6
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
    Address     As String   'Patient의 Address 와 Street를 합친것
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

Public Type Dept           ' 부서 전체(C,T,E)
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
Public gCDept As CDept      '진료과 정보 Dept의 Gubun 이 'C' 인것

Public Type TxDept
    Code        As String
    EName       As String
    KName       As String
    Alias       As String
End Type
Public gTxDept As TxDept      '진료지원부서 정보 Dept의 Gubun 이 'T' 인것

Public Type EDept
    Code        As String
    EName       As String
    KName       As String
    Alias       As String
End Type
Public gEDept As EDept      '기타 부서 정보 Dept의 Gubun 이 'E' 인것


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
Public gUser As User        '프로그램 사용자 정보

Public Type Slip
    Code        As String
    Name        As String
    Alias       As String
End Type
Public gSlip As Slip        'SlipClass 정보

Public Type LabMaster
    Code        As String
    Name        As String
    RoomNo      As String
    Phone       As String
End Type
Public gLab As LabMaster          '임상병리과 파트 정보


Public Type CaseMaster
    Code        As String
    Name        As String
    Remark      As String
End Type
Public gCase As CaseMaster      ' 용기마스터

Public Type UnitMaster
    Code        As String
    Name        As String
End Type
Public gUnit As UnitMaster      ' 단위마스터

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
Public gExam As Exam        '검사 항목 정보

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
Public gTrust As Trust      '거래처 정보

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
Public gEquip As Equip      ' 장비

Public Type Ward
    Code        As String
    Room        As String
    Name        As String
    BedCheck    As String
    Nurse       As String
    UseFlag     As String
    Remark      As String
End Type
Public gWard As Ward        '병동


    
    


