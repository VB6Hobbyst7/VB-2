VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsS2DSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event UnloadForm()

Private MyAbout As frmAbout
Private WithEvents MyDatabase As frmDatabase
Attribute MyDatabase.VB_VarHelpID = -1
Private WithEvents MyBuilding As frmBuilding
Attribute MyBuilding.VB_VarHelpID = -1
Private MyRegEdit As frmRegEdit

Private mvarAppName As String
Private mvarAppPath As String
Private mvarExeName As String
Private mvarTradeMark As String

Private mvarEmpId As String
Private mvarEmpNm As String
Private mvarEmpSNm As String
Private mvarSex As String
Private mvarDOB As String
Private mvarSSN As String
Private mvarShiftCd As String
Private mvarLogonId As String
Private mvarPassword As String
Private mvarDeptCd As String
Private mvarHireDt As String
Private mvarEntDt As String
Private mvarFireDt As String
Private mvarDegree As String
Private mvarTechFg As String
Private mvarTelNo As String
Private mvarCellNo As String
Private mvarNote As String

Private mvarProjectId As String         ' ������Ʈ��
Private mvarVersion As String           ' ����
Private mvarCopyright As String         ' ���۱�
Private mvarHospital As String          ' ����
Private mvarHelpLine As String          ' ����� ��ȭ��ȣ

Private mvarUseBuildingInfo As String   '-- �ǹ����� �̿뿩��
Private mvarBuildingCd As String        '-- �ǹ��ڵ�
Private mvarBuildingNm As String        '-- �ǹ���
Private mvarBuildingNo As Integer       '-- �ǹ���ȣ
Private mvarBuildingList As String      '-- File Server IP
Private mvarFileServer As String        '-- File Server IP

Private mvarServerNm As String          '-- ������
Private mvarDatabaseNm As String        '-- ����Ÿ���̽���
Private mvarDBLoginId As String         '-- �α�Ƶ�
Private mvarDBPassword As String        '-- �н�����
Private mvarDBType As Long            '-- ����Ÿ���̽� ����
Private mvarServerRegistered As Boolean
Private mvarShowAtStartup As String
Private mvarRunSplash As String

Private mvarSB_ConnStatus As Integer

Private mvarDownloadFg As Boolean

'frmDatabase / frmBuilding
Private mvarCancelIsEnd As Boolean
Private mvarButtonCheck As String
Private mvarRegCanceled As Boolean
Private mvarRegChanged As Boolean

'[�Ӽ�] - Project Id
Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property
Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

'[�Ӽ�] - Version
Public Property Let Version(ByVal vData As String)
    mvarVersion = vData
End Property
Public Property Get Version() As String
    Version = mvarVersion
End Property

'[�Ӽ�] - Copyright
Public Property Let Copyright(ByVal vData As String)
    mvarCopyright = vData
End Property
Public Property Get Copyright() As String
    Copyright = mvarCopyright
End Property

'[�Ӽ�] - ����
Public Property Let Hospital(ByVal vData As String)
    mvarHospital = vData
End Property
Public Property Get Hospital() As String
    Hospital = mvarHospital
End Property

'[�Ӽ�] - ����� ��ȭ��ȣ
Public Property Let HelpLine(ByVal vData As String)
    mvarHelpLine = vData
End Property
Public Property Get HelpLine() As String
    HelpLine = mvarHelpLine
End Property

'[�Ӽ�] - Registry ��� ��ҿ���
Public Property Get RegCanceled() As String
    RegCanceled = mvarRegCanceled
End Property
Public Property Get RegChanged() As String
    RegChanged = mvarRegChanged
End Property

'[�Ӽ�] - Application ��
Public Property Let AppName(ByVal vData As String)
    mvarAppName = vData
End Property
Public Property Get AppName() As String
    AppName = mvarAppName
End Property

'[�Ӽ�] - Trade Maek (��ǥ)
Public Property Let TradeMark(ByVal vData As String)
    mvarTradeMark = vData
End Property
Public Property Get TradeMark() As String
    TradeMark = mvarTradeMark
End Property

'[�Ӽ�] - �ǹ����� ��뿩��
Public Property Let UseBuildingInfo(ByVal vData As String)
    mvarUseBuildingInfo = vData
End Property
Public Property Get UseBuildingInfo() As String
    UseBuildingInfo = mvarUseBuildingInfo
End Property

'[�Ӽ�] - BuildingList
Public Property Let BuildingList(ByVal vData As String)
    mvarBuildingList = vData
End Property
Public Property Get BuildingList() As String
    BuildingList = mvarBuildingList
End Property

'[�Ӽ�] - Download ���� ����
Public Property Let DownloadFg(ByVal vData As Boolean)
    mvarDownloadFg = vData
End Property
Public Property Get DownloadFg() As Boolean
    DownloadFg = mvarDownloadFg
End Property

'[�Ӽ�] - DB ��й�ȣ
Public Property Let DBPassword(ByVal vData As String)
    mvarDBPassword = vData
End Property
Public Property Get DBPassword() As String
    DBPassword = mvarDBPassword
End Property

'[�Ӽ�] - DB ����
Public Property Let DbTYPE(ByVal vData As Long)
    mvarDBType = vData
End Property

Public Property Get DbTYPE() As Long
    DbTYPE = mvarDBType
End Property

'[�Ӽ�] - DB �����ID
Public Property Let DBLoginId(ByVal vData As String)
    mvarDBLoginId = vData
End Property
Public Property Get DBLoginId() As String
    DBLoginId = mvarDBLoginId
End Property

'[�Ӽ�] - DB ��
Public Property Let DatabaseNm(ByVal vData As String)
    mvarDatabaseNm = vData
End Property
Public Property Get DatabaseNm() As String
    DatabaseNm = mvarDatabaseNm
End Property

'[�Ӽ�] - ������
Public Property Let ServerNm(ByVal vData As String)
    mvarServerNm = vData
End Property
Public Property Get ServerNm() As String
    ServerNm = mvarServerNm
End Property

'[�Ӽ�] - ���ϼ��� IP
Public Property Let FileServer(ByVal vData As String)
    mvarFileServer = vData
End Property
Public Property Get FileServer() As String
    FileServer = mvarFileServer
End Property

'[�Ӽ�] - �ǹ���ȣ
Public Property Let BuildingNo(ByVal vData As String)
    mvarBuildingNo = vData
End Property
Public Property Get BuildingNo() As String
    BuildingNo = mvarBuildingNo
End Property

'[�Ӽ�] - �ǹ���
Public Property Let BuildingNm(ByVal vData As String)
    mvarBuildingNm = vData
End Property
Public Property Get BuildingNm() As String
    BuildingNm = mvarBuildingNm
End Property

'[�Ӽ�] - �ǹ��ڵ�
Public Property Let BuildingCd(ByVal vData As String)
    mvarBuildingCd = vData
End Property
Public Property Get BuildingCd() As String
    BuildingCd = mvarBuildingCd
End Property

'[�Ӽ�] - ���� Remark
Public Property Let Note(ByVal vData As String)
    mvarNote = vData
End Property
Public Property Get Note() As String
    Note = mvarNote
End Property

'[�Ӽ�] - ���� �޴�����ȣ
Public Property Let CellNo(ByVal vData As String)
    mvarCellNo = vData
End Property
Public Property Get CellNo() As String
    CellNo = mvarCellNo
End Property

'[�Ӽ�] - ���� ��ȭ��ȣ
Public Property Let TelNo(ByVal vData As String)
    mvarTelNo = vData
End Property
Public Property Get TelNo() As String
    TelNo = mvarTelNo
End Property

'[�Ӽ�] - ���� ��翩��
Public Property Let TechFg(ByVal vData As String)
    mvarTechFg = vData
End Property
Public Property Get TechFg() As String
    TechFg = mvarTechFg
End Property

'[�Ӽ�] - ���� ����
Public Property Let Degree(ByVal vData As String)
    mvarDegree = vData
End Property
Public Property Get Degree() As String
    Degree = mvarDegree
End Property

'[�Ӽ�] - ���� �����
Public Property Let FireDt(ByVal vData As String)
    mvarFireDt = vData
End Property
Public Property Get FireDt() As String
    FireDt = mvarFireDt
End Property

'[�Ӽ�] - �������� �����
Public Property Let EntDt(ByVal vData As String)
    mvarEntDt = vData
End Property
Public Property Get EntDt() As String
    EntDt = mvarEntDt
End Property

'[�Ӽ�] - ���� �Ի���
Public Property Let HireDt(ByVal vData As String)
    mvarHireDt = vData
End Property
Public Property Get HireDt() As String
    HireDt = mvarHireDt
End Property

'[�Ӽ�] - ���� �μ��ڵ�
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property
Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

'[�Ӽ�] - ���α׷� Login Password
Public Property Let Password(ByVal vData As String)
    mvarPassword = vData
End Property
Public Property Get Password() As String
    Password = mvarPassword
End Property

'[�Ӽ�] - ���α׷� Login ID
Public Property Let LogonId(ByVal vData As String)
    mvarLogonId = vData
End Property
Public Property Get LogonId() As String
    LogonId = mvarLogonId
End Property

'[�Ӽ�] - ���� �ٹ�����
Public Property Let ShiftCd(ByVal vData As String)
    mvarShiftCd = vData
End Property
Public Property Get ShiftCd() As String
    ShiftCd = mvarShiftCd
End Property

'[�Ӽ�] - ���� �ֹι�ȣ
Public Property Let SSN(ByVal vData As String)
    mvarSSN = vData
End Property
Public Property Get SSN() As String
    SSN = mvarSSN
End Property

'[�Ӽ�] - �����������
Public Property Let DOB(ByVal vData As String)
    mvarDOB = vData
End Property
Public Property Get DOB() As String
    DOB = mvarDOB
End Property

'[�Ӽ�] - ��������
Public Property Let Sex(ByVal vData As String)
    mvarSex = vData
End Property
Public Property Get Sex() As String
    Sex = mvarSex
End Property

'[�Ӽ�] - ������(Short)
Public Property Let EmpSNm(ByVal vData As String)
    mvarEmpSNm = vData
End Property
Public Property Get EmpSNm() As String
    EmpSNm = mvarEmpSNm
End Property

'[�Ӽ�] - ������(Logn)
Public Property Let EmpNm(ByVal vData As String)
    mvarEmpNm = vData
End Property
Public Property Get EmpNm() As String
    EmpNm = mvarEmpNm
End Property

'[�Ӽ�] - ����ID
Public Property Let EmpId(ByVal vData As String)
    mvarEmpId = vData
End Property
Public Property Get EmpId() As String
    EmpId = mvarEmpId
End Property

'[�Ӽ�] - Server���� Registry��� ����
Public Property Let ServerRegistered(ByVal vData As Boolean)
    mvarServerRegistered = vData
End Property
Public Property Get ServerRegistered() As Boolean
    ServerRegistered = mvarServerRegistered
End Property

'[�Ӽ�] - ���۽� �������׺��� ����
Public Property Let ShowAtStartup(ByVal vData As String)
    mvarShowAtStartup = vData
End Property
Public Property Get ShowAtStartup() As String
    ShowAtStartup = mvarShowAtStartup
End Property

'[�Ӽ�] - ���۽� Splash ȭ�� ����
Public Property Let RunSplash(ByVal vData As String)
    mvarRunSplash = vData
End Property
Public Property Get RunSplash() As String
    RunSplash = mvarRunSplash
End Property

'[�Ӽ�] - CancelIsEnd
Public Property Let CancelIsEnd(ByVal vData As Boolean)
    mvarCancelIsEnd = vData
End Property
Public Property Get CancelIsEnd() As Boolean
    CancelIsEnd = mvarCancelIsEnd
End Property

'[�Ӽ�] - Button Check
Public Property Let ButtonCheck(ByVal vData As String)
    mvarButtonCheck = vData
End Property
Public Property Get ButtonCheck() As String
    ButtonCheck = mvarButtonCheck
End Property

'[�Ӽ�] - Button Check
Public Property Let AppPath(ByVal vData As String)
    mvarAppPath = vData
End Property
Public Property Get AppPath() As String
    AppPath = mvarAppPath
End Property

'[�Ӽ�] - Button Check
Public Property Let ExeName(ByVal vData As String)
    mvarExeName = vData
End Property
Public Property Get ExeName() As String
    ExeName = mvarExeName
End Property

'[Method] - Database Configuration â ����
Public Sub LoadDatabaseConfig()

    mvarRegChanged = False
    mvarRegCanceled = False
    
    Set MyDatabase = frmDatabase
    With MyDatabase
        Call .ApplyButton(mvarButtonCheck)
        Call medAlwaysOn(MyDatabase, 1)
        .WindowState = vbNormal
        
        .OldServerNm = mvarServerNm
        .OldDatabaseNm = mvarDatabaseNm
        .OldLoginId = mvarDBLoginId
        .OldPassword = mvarDBPassword
        .OldDbType = mvarDBType
        .SetServerInfo
        .Show vbModal
    End With
    
End Sub

'[Method] - �ǹ����� ���� â ����
Public Sub LoadBuildingInfo()
    
    Set MyBuilding = frmBuilding
    With MyBuilding
        .CurBuildingCd = mvarBuildingCd
        .CurBuildingNo = mvarBuildingNo
        .CurBuildingNm = mvarBuildingNm
        .BuildingList = mvarBuildingList
        .WindowState = vbNormal
        Call medAlwaysOn(MyBuilding, 1)
        Call .ApplyButton(mvarButtonCheck)
        Call .LoadBuildingList
        .Show vbModal
    End With
End Sub

'[Method] - Aboutȭ�� ����
Public Sub LoadAbout(Optional ByVal blnEnd As Boolean = False)

    Set MyAbout = frmAbout
    Call medAlwaysOn(MyAbout, 1)
    With MyAbout
        .ProductName = mvarProjectId
        .Version = mvarVersion
        .Copyright = mvarCopyright
        .Hospital = mvarHospital
        .WindowState = vbNormal
        .SetValues
        
        .cmdOK.Visible = Not blnEnd
        
        If blnEnd Then
            .Show
        Else
            .Show 1
        End If
    
    End With
    
End Sub

'[Method] - Registry��� â ����
Public Sub LoadRegEdit()
    
    Set MyRegEdit = frmRegEdit
    Call medAlwaysOn(MyRegEdit, 1)
    MyRegEdit.TradeMark = mvarTradeMark
    MyRegEdit.WindowState = vbNormal
    MyRegEdit.Show vbModal
    
End Sub

'[Method] - Application �� Set (Registry ��� Key : ��- Schweitzer LIS)
Public Sub SetAppName(ByVal pValue As String)
    mvarAppName = pValue
    RegAppName = pValue
End Sub

'[Method] - Registry ���� �о����
Public Sub ReadRegistryInfo()
    
    '����, ����Ÿ���̽�, �α��ξƵ�, ��й�ȣ
    mvarServerNm = GetSetting(AppName:=mvarAppName, Section:=RegSsSvr, Key:=RegK1Svr, Default:="")
    mvarDatabaseNm = GetSetting(AppName:=mvarAppName, Section:=RegSsSvr, Key:=RegK2Svr, Default:="")
    mvarDBLoginId = GetSetting(AppName:=mvarAppName, Section:=RegSsSvr, Key:=RegK3Svr, Default:="")
    mvarDBPassword = GetSetting(AppName:=mvarAppName, Section:=RegSsSvr, Key:=RegK4Svr, Default:="")
    mvarDBType = GetSetting(AppName:=mvarAppName, Section:=RegSsSvr, Key:=RegK5Svr, Default:=0)
    
    DCM_DbType = mvarDBType
    
    If mvarServerNm = "" Or mvarDatabaseNm = "" Then
        mvarServerNm = "(Undefined)"
        mvarServerRegistered = False
    Else
        mvarServerRegistered = True
    End If
    '������
    mvarHospital = GetSetting(AppName:=mvarAppName, Section:=RegSsSet, Key:=RegK2Set, Default:="")
    mvarHelpLine = GetSetting(AppName:=mvarAppName, Section:=RegSsSet, Key:=RegK3Set, Default:="")
    '�ǹ����� ��뿩��
    mvarUseBuildingInfo = GetSetting(AppName:=mvarAppName, Section:=RegSsBld, Key:=RegK0Bld, Default:="0")
    '�ǹ�����(�ڵ�,��,��ȣ)
    If mvarUseBuildingInfo = 0 Then
        mvarBuildingCd = "10"
        mvarBuildingNm = "����"
        mvarBuildingNo = 1
    Else
        mvarBuildingCd = GetSetting(AppName:=mvarAppName, Section:=RegSsBld, Key:=RegK1Bld, Default:="")
        mvarBuildingNm = GetSetting(AppName:=mvarAppName, Section:=RegSsBld, Key:=RegK2Bld, Default:="(�ǹ����� ����)")
        mvarBuildingNo = GetSetting(AppName:=mvarAppName, Section:=RegSsBld, Key:=RegK3Bld, Default:=0)
    End If
    
    '�������� ��ȸ����
    mvarShowAtStartup = GetSetting(mvarAppName, RegSsOpt, RegK1Opt, "1")
    mvarRunSplash = GetSetting(mvarAppName, RegSsOpt, RegK2Opt, "1")

End Sub

'[Method] - Registry ��� : Application Path
Public Sub CheckAppPath(ByVal strCurAppPath As String)

    Dim strOldAppPath As String

    strOldAppPath = GetSetting(AppName:=mvarAppName, Section:=RegSsApp, Key:=RegK1App, Default:="")
    If strOldAppPath <> strCurAppPath Then SaveSetting mvarAppName, RegSsApp, RegK1App, strCurAppPath
    mvarAppPath = strCurAppPath
    mvarExeName = GetSetting(AppName:=mvarAppName, Section:=RegSsApp, Key:=RegK2App, Default:="")

End Sub

'[Method] - Registry ��� : File Server IP & Path
Public Sub SetFileServer(ByVal strServerIP As String)
    SaveSetting mvarAppName, RegSsSet, RegK1Set, strServerIP
End Sub

'[Method] - File Server�� IP �� Path�� �����´�.
Public Function GetFileServer(ByVal sIndexCd As String, ByVal sProjectId As String) As String
    
    GetFileServer = " select text1 as svrpath, field3 as version from " & T_COM003 & _
                    " where cdindex = '" & sIndexCd & "' and cdval1 = '" & sProjectId & "' and field1 = '1' "
    
End Function

Private Sub Class_Terminate()
    If Not (MyAbout Is Nothing) Then
        Unload MyAbout
        Set MyAbout = Nothing
    End If
End Sub

'[Event] - Building ������ ����Ǿ��� ��� �߻�.
Private Sub MyBuilding_SetBuilding()
    
    mvarBuildingCd = MyBuilding.CurBuildingCd
    mvarBuildingNm = MyBuilding.CurBuildingNm
    mvarBuildingNo = MyBuilding.CurBuildingNo

End Sub

'[Event] - Database��� â���� Cancel��ư�� Click�� �߻�
Private Sub MyDatabase_RegCanceled()
    mvarRegCanceled = True
End Sub

'[Event] - Database��� â���� Save��ư�� Click�� �߻�
Private Sub MyDatabase_RegSaved()
    
    With MyDatabase
        mvarServerNm = .txtServer.Text
        mvarDatabaseNm = .txtDbNm.Text
        mvarDBLoginId = .txtLogin.Text
        mvarDBPassword = .txtPwd.Text
        mvarDBType = .cboDBType.ListIndex
        mvarRegChanged = .RegChanged
    End With

End Sub