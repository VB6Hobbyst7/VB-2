Attribute VB_Name = "Hi_Code"
Option Explicit

' *=====================================================*
' *           Gloval Variable Define                    *
' *=====================================================*
Global Const MAX_NUM = 30        ' 최대등록수

Type TEST_NAME_TABLE             ' 검사항목명 Table
'    SeqNo   As String
    EqCd    As String * 2
    Name    As String '* 15
    Code    As String
End Type
Global TestNameTable(MAX_NUM) As TEST_NAME_TABLE

Type TEST_RESULT_TABLE           ' 결과용 Table
    SeqNo   As String
    TestNo  As String '* 2
    Result  As String * 9
End Type
Global TestResultTable(MAX_NUM) As TEST_RESULT_TABLE

Type TEST_HEADER_TABLE            ' Sample ID Table
     RackNo As String * 4         ' Rack No.
     SampNo As String * 5         ' Sample No.
     PosiNo As String * 1         ' Posotion No.
     Lab_ID As String * 14        ' ID No.
End Type
Global TestHeaderTable As TEST_HEADER_TABLE

Global strLAB_PATH As String    'Interface 디렉토리 저장
Global ResDb_Path  As String    '결과 DB의 위치 저장

Global DbComm   As Database
Global TbComm   As Recordset
Global Db       As Database
Global TableCd  As Recordset
Global TableVar As Recordset
Global TableSam As Recordset
Global DbResult   As Database
Global TbSample   As Recordset
Global TbResult   As Recordset

Type COMM_CONFIG
    ComPort       As Integer
    BaudRate      As String
    DataBit       As String
    StopBit       As String
    Parity        As String
End Type
Global CommCfg    As COMM_CONFIG

'Global DbCm        As Database
'Global TableVar    As Table
Global tlFileName  As String
'Global RegFrom     As String
'Global RegUntil    As String
'Global RegExit     As Integer

' *============ Debug ============*
Global TestDat  As String * 60
Global TestPtr     As Integer
'Global StartSampleNo As Integer

' *========== Common ===========*
Global dDate   As String        'Data 저장기간

Global SqlDoc   As String
Global ChkRow   As Integer
Global sStr As String

Global spdRow   As Integer
Global DelFileChk   As Integer
Global StartSampleNo         As Integer

Public FrmTag   As Integer


'
'   Hitachi DB 생성
'
Public Sub Create_HIDB(fnstr As String)
    
    Dim Db As Database
    
    Dim NewTbd  As New TableDef
    Dim NewIdx  As New Index
    Dim NewIdx1 As New Index
    Dim SampNo  As New Field
    Dim EqCode  As New Field
    Dim SeqNo   As New Field
    Dim Result  As New Field
       
    Dim FileName As String
    
    On Error GoTo RepairDB
    
    FileName = strLAB_PATH & "data\"
    If iFileExists(FileName + fnstr + ".mdb") Then
        Set Db = OpenDatabase(FileName + fnstr + ".mdb")
        Exit Sub
    End If
    
    Set Db = CreateDatabase(FileName + fnstr + ".mdb", DB_LANG_GENERAL)
    
    NewTbd.Name = "Hi_Result"
    
    SampNo.Name = "SampNo"
    SampNo.Type = DB_TEXT
    SampNo.Size = 5
    NewTbd.Fields.Append SampNo
    
    EqCode.Name = "EqCode"
    EqCode.Type = DB_TEXT
    EqCode.Size = 2
    NewTbd.Fields.Append EqCode
    
    SeqNo.Name = "SeqNo"
    SeqNo.Type = DB_TEXT
    SeqNo.Size = 2
    NewTbd.Fields.Append SeqNo
    
    Result.Name = "Result"
    Result.Type = DB_TEXT
    Result.Size = 9
    Result.AllowZeroLength = True
    NewTbd.Fields.Append Result
    
    NewIdx.Name = "PrimaryKey"
    NewIdx.Fields = "SampNo;EqCode"
    NewIdx.Primary = True
    NewTbd.Indexes.Append NewIdx
    
    NewIdx1.Name = "SampNo_Idx"
    NewIdx1.Fields = "SampNo"
    NewIdx1.Primary = False
    NewTbd.Indexes.Append NewIdx1
    
    Db.TableDefs.Append NewTbd
    '============
    
    Dim NewTbd2  As New TableDef
    Dim NewIdx2  As New Index
    Dim NewIdx3  As New Index
    Dim Samp_No  As New Field
    Dim Lab_ID   As New Field
    Dim RackNo  As New Field
    Dim PosNo   As New Field
    Dim RegChk  As New Field    'Server등록여부
    Dim PrtChk  As New Field    '출력여부
    
    NewTbd2.Name = "Hi_Sample"
    
    Samp_No.Name = "SampNo"
    Samp_No.Type = DB_TEXT
    Samp_No.Size = 5
    NewTbd2.Fields.Append Samp_No
    
    Lab_ID.Name = "Lab_ID"
    Lab_ID.Type = DB_TEXT
    Lab_ID.Size = 14
    NewTbd2.Fields.Append Lab_ID
    
    RackNo.Name = "RackNo"
    RackNo.Type = DB_TEXT
    RackNo.Size = 4
    NewTbd2.Fields.Append RackNo

    PosNo.Name = "PosNo"
    PosNo.Type = DB_TEXT
    PosNo.Size = 1
    NewTbd2.Fields.Append PosNo
        
    RegChk.Name = "RegChk"
    RegChk.Type = DB_TEXT
    RegChk.Size = 1
    NewTbd2.Fields.Append RegChk
    
    PrtChk.Name = "PrtChk"
    PrtChk.Type = DB_TEXT
    PrtChk.Size = 1
    NewTbd2.Fields.Append PrtChk
    
    NewIdx2.Name = "PrimaryKey"
    NewIdx2.Fields = "SampNo"
    NewIdx2.Primary = True
    NewTbd2.Indexes.Append NewIdx2
    
    NewIdx3.Name = "Lab_ID_Idx"
    NewIdx3.Fields = "Lab_ID"
    NewIdx3.Primary = False
    NewTbd2.Indexes.Append NewIdx3
    
    Db.TableDefs.Append NewTbd2
    '============
    
    Db.Close
    Exit Sub
    
RepairDB:
    If Err = 3049 Then
        MsgBox "데이타가 손상되어 있습니다. 확인을 누르시면 데이타를 복구합니다.", _
                vbInformation
        RepairDatabase (FileName + fnstr + ".mdb")
        Set Db = OpenDatabase(FileName + fnstr + ".mdb")
    End If
    Resume Next
    
End Sub

