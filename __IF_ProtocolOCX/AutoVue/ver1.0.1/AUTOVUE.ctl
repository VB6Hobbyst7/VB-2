VERSION 5.00
Begin VB.UserControl AUTOVUE 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2100
   LockControls    =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   2100
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   90
      Top             =   630
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "AutoVue"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   1500
   End
End
Attribute VB_Name = "AUTOVUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sRegNo = 0
Const m_def_sOrderFormat = 0
Const m_def_p_sFormula = "0"
'Const m_def_DBPath = 0
Const m_def_Settings = ""
Const m_def_sRstFileNm = "0"
Const m_def_sRstFilePath = "0"
Const m_def_sVersion = "0"
Const m_def_EqName = "0"
Const m_def_bUseBarcode = 0
Const m_def_iPhase = 0
Const m_def_iSendPhase = 0
Const m_def_sTestMode = "0"
Const m_def_iFrameN = 0
Const m_def_p_sID = "0"
Const m_def_p_sSeq = "0"
Const m_def_p_sRack = "0"
Const m_def_p_sPos = "0"
Const m_def_p_iOrdCnt = 0
Const m_def_p_sTIFCd = "0"
Const m_def_OpenPW = "0"
Const m_def_EditPW = "0"
'속성 변수:
Dim m_p_sRegNo As Variant
Dim m_sOrderFormat As Variant
Dim m_p_sFormula As String
'Dim m_DBPath As Variant
Dim m_Settings As String
Dim m_sRstFileNm As String
Dim m_sRstFilePath As String
Dim m_sVersion As String
Dim m_EqName As String
Dim m_bUseBarcode As Boolean
Dim m_iPhase As Integer
Dim m_iSendPhase As Integer
Dim m_sTestMode As String
Dim m_iFrameN As Integer
Dim m_p_sID As String
Dim m_p_sSeq As String
Dim m_p_sRack As String
Dim m_p_sPos As String
Dim m_p_iOrdCnt As Integer
Dim m_p_sTIFCd As String
Dim m_OpenPW As String
Dim m_EditPW As String
'이벤트 선언:
Event RequestData(sUploadPath$, sDownLoadPath$)
'Event RequestData(sUploadPath$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTifcd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sOther1$)
Event RequestCurOrder(sID$)
Event SendOrderOK(sID$)
'Event GetFormula(sTifcd$)
Event ChkCurPatID(sPatID$, sRetCd$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)


'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim wkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim iSpaceCnt   As Integer

'for AlloScan
Private Type PANELCODEINFO
    FieldNm(100)    As String
    Code(100)       As String
End Type
Private tPnlCdInfo()    As PANELCODEINFO

'
'   결과정보 구조체 초기화
'
Private Sub Init_pResultInfo()
    
    With pResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .KIND = ""
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
End Sub

Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .SEQNO = m_p_sSeq
            .RACK = m_p_sRack
            .POS = m_p_sPos
            .ORDCNT = 0
            .OTHER = m_p_sRegNo
        End With
        Exit Sub
    End If
    
    ReDim tmpData(m_p_iOrdCnt) As String
    tmpData() = Split(m_p_sTIFCd, Chr(124))
    
    With pSampleInfo
        .ID = m_p_sID
        .SEQNO = m_p_sSeq
        .RACK = m_p_sRack
        .POS = m_p_sPos
        .ORDCNT = m_p_iOrdCnt
        .OTHER = m_p_sRegNo
        
        ReDim .IFCD(.ORDCNT)
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
    End With
        
End Sub


Private Sub Timer1_Timer()
    On Error GoTo ErrRtn

    Dim FSO As FileSystemObject
    Dim objFile As Object
    Dim sUplFilePath As String
    Dim sDnlFilePath As String
    Dim sRtnVal As String
    Dim sOneRow() As String
    Dim tmpField()  As String
    Dim tmpData() As String
    Dim tmpBarcd As String
    Dim sRecType As String
    Dim sOrderFile As String
    Dim iPatientCnt As Integer
    Dim i%, j%
    Dim sFileNm As String
    Dim tmpSeqNo$
    Dim tmpIFCd$, tmpRst$, tmpUnit$

    Timer1.Enabled = False

    RaiseEvent RequestData(sUplFilePath, sDnlFilePath)

    Set FSO = New FileSystemObject

    If FSO.FolderExists(sUplFilePath) Then
        For Each objFile In FSO.GetFolder(sUplFilePath).Files
            If UCase(Right(objFile.Name, 3)) = UCase(sOrderFormat) Then
                sRtnVal = ""
                Open sUplFilePath & "\" & objFile.Name For Input Shared As #66
                    Do While Not EOF(66)
                        sRtnVal = sRtnVal & Input(1, #66)
                    Loop
                Close #66

                sOneRow = Split(sRtnVal, vbCrLf)

                iPatientCnt = 0
                sOrderFile = ""
'                sOrderFile = sOrderFile & "H|\^&|||LIS|||||||P|1|" & Format(Now, "YYYYMMDDHHmmdd") & vbCrLf
                sOrderFile = sOrderFile & "H|\^&||||||||||||" & Format(Now, "YYYYMMDDHHmmdd") & vbCrLf

                For i = 0 To UBound(sOneRow) - 1
                    If Trim(sOneRow(i)) = "" Then Exit For

                    sRecType = Left(sOneRow(i), 1)

                    Select Case sRecType
                        Case "H"
                        Case "Q"
                            tmpField = Split(sOneRow(i), Chr(124))
                            tmpData = Split(tmpField(2), "^")
                            tmpBarcd = Trim(tmpData(1))

                            RaiseEvent RequestCurOrder(tmpBarcd)

                            Call Get_OrderString

                            If m_p_iOrdCnt > 0 Then
                                iPatientCnt = iPatientCnt + 1
                                sOrderFile = sOrderFile & "P|" & CStr(iPatientCnt) & "|" & pSampleInfo.OTHER & "||^^" & pSampleInfo.OTHER & "|" & pSampleInfo.OTHER & "^|||M||||||||||||||||||||||||||" & vbCrLf
                                For j = 1 To pSampleInfo.ORDCNT
                                    sOrderFile = sOrderFile & "O|" & CStr(j) & "|" & tmpBarcd & "||" & pSampleInfo.IFCD(j) & "|N|" & Format(Now, "YYYYMMDDHHmmdd") & "|||||||||CENTBLOOD|||||||" & Format(Now, "YYYYMMDDHHmmdd") & "|||F|||||" & vbCrLf
                                Next j
                                
                                RaiseEvent SendOrderOK(tmpBarcd)
'                            Else
'                                iPatientCnt = iPatientCnt + 1
'                                sOrderFile = sOrderFile & "P|" & CStr(iPatientCnt) & "|" & pSampleInfo.OTHER & "||^^" & pSampleInfo.OTHER & "|" & pSampleInfo.OTHER & "^|||M||||||||||||||||||||||||||" & vbCrLf
                            End If

                    End Select
                Next i

                sOrderFile = sOrderFile & "L" & vbCrLf
                
                If sTestMode = "77" Then
                    RaiseEvent PrintSendLog(sOrderFile)
                End If

                Open sDnlFilePath & "\" & Mid(objFile.Name, 1, InStr(objFile.Name, ".")) & "dnl" For Output Shared As #67
                Print #67, sOrderFile
                Close #67
                
                FSO.CopyFile sUplFilePath & "\" & objFile.Name, sUplFilePath & "\processed\" & objFile.Name
                FSO.DeleteFile sUplFilePath & "\" & objFile.Name
            ElseIf UCase(Right(objFile.Name, 3)) = "UPL" Then
                Open sUplFilePath & "\" & objFile.Name For Input Shared As #77
                    sRtnVal = ""
                    Do While Not EOF(77)
                        sRtnVal = sRtnVal & Input(1, #77)
                    Loop
                Close #77
                
                If sTestMode = "77" Then
                    RaiseEvent PrintRcvLog(sRtnVal)
                End If
                
                sOneRow = Split(sRtnVal, vbCrLf)
                
                For i = 0 To UBound(sOneRow) - 1
                    If Trim(sOneRow(i)) = "" Then Exit For
                    
                    sRecType = Left(sOneRow(i), 1)
                    
                    Select Case sRecType
                        Case "H"
                        Case "P"
                        Case "O"
                            '결과값 등록/화면 표시 처리...
                            With pResultInfo
                                If .RSTCNT > 0 Then
                                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .OTHER)
                                End If
                            End With
                            
                            tmpBarcd = "": tmpIFCd = "": tmpRst = ""
                            Call Init_pResultInfo
                            
                            tmpField = Split(sOneRow(i), Chr(124))
                            tmpBarcd = Trim(tmpField(2))
                            pResultInfo.ID = tmpBarcd
                        
                        Case "R"
                            tmpField = Split(sOneRow(i), Chr(124))
                            tmpIFCd = Trim(tmpField(2))
                            tmpRst = Trim(tmpField(3))
                            
                            '결과정보 구조체에 저장
                            With pResultInfo
                                '결과값 누적
                                .RSTCNT = .RSTCNT + 1
                                .IFCD = .IFCD & tmpIFCd & Chr(124)
                                .RST1 = .RST1 & tmpRst & Chr(124)
                                .RST2 = .RST2 & Chr(124)
                                .UNIT = .UNIT & Chr(124)
                                .FLAG = .FLAG & Chr(124)
                            End With
                            
                        Case "M"
                            tmpField = Split(sOneRow(i), Chr(124))
                            tmpIFCd = Trim(tmpField(2))
                            tmpRst = Trim(tmpField(5))
                            tmpRst = Split(tmpRst, "^")(0)
                            
                            '결과정보 구조체에 저장
                            With pResultInfo
                                '결과값 누적
                                .RSTCNT = .RSTCNT + 1
                                .IFCD = .IFCD & tmpIFCd & Chr(124)
                                .RST1 = .RST1 & tmpRst & Chr(124)
                                .RST2 = .RST2 & Chr(124)
                                .UNIT = .UNIT & Chr(124)
                                .FLAG = .FLAG & Chr(124)
                            End With
                        
                        Case "L"
                            '결과값 등록/화면 표시 처리...
                            With pResultInfo
                                If .RSTCNT > 0 Then
                                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .OTHER)
                                End If
                            End With
                            
                            tmpBarcd = "": tmpIFCd = "": tmpRst = ""
                            Call Init_pResultInfo
                        
                    End Select
                Next i
                
                FSO.CopyFile sUplFilePath & "\" & objFile.Name, sUplFilePath & "\processed\" & objFile.Name
                FSO.DeleteFile sUplFilePath & "\" & objFile.Name
            End If
        Next
    End If
    
    Set FSO = Nothing

    Timer1.Enabled = True

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Timer1_Timer 오류발생 - " & Err.Description)
        Timer1.Enabled = True
    End If
End Sub

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_OpenPW = PropBag.ReadProperty("OpenPW", m_def_OpenPW)
    m_EditPW = PropBag.ReadProperty("EditPW", m_def_EditPW)
    m_EqName = PropBag.ReadProperty("EqName", m_def_EqName)
    m_bUseBarcode = PropBag.ReadProperty("bUseBarcode", m_def_bUseBarcode)
    m_iPhase = PropBag.ReadProperty("iPhase", m_def_iPhase)
    m_iSendPhase = PropBag.ReadProperty("iSendPhase", m_def_iSendPhase)
    m_sTestMode = PropBag.ReadProperty("sTestMode", m_def_sTestMode)
    m_iFrameN = PropBag.ReadProperty("iFrameN", m_def_iFrameN)
    m_p_sID = PropBag.ReadProperty("p_sID", m_def_p_sID)
    m_p_sSeq = PropBag.ReadProperty("p_sSeq", m_def_p_sSeq)
    m_p_sRack = PropBag.ReadProperty("p_sRack", m_def_p_sRack)
    m_p_sPos = PropBag.ReadProperty("p_sPos", m_def_p_sPos)
    m_p_iOrdCnt = PropBag.ReadProperty("p_iOrdCnt", m_def_p_iOrdCnt)
    m_p_sTIFCd = PropBag.ReadProperty("p_sTIFCd", m_def_p_sTIFCd)
    m_sVersion = PropBag.ReadProperty("sVersion", m_def_sVersion)
    m_sRstFileNm = PropBag.ReadProperty("sRstFileNm", m_def_sRstFileNm)
    m_sRstFilePath = PropBag.ReadProperty("sRstFilePath", m_def_sRstFilePath)
    m_Settings = PropBag.ReadProperty("Settings", m_def_Settings)

    m_p_sFormula = PropBag.ReadProperty("p_sFormula", m_def_p_sFormula)
    m_sOrderFormat = PropBag.ReadProperty("sOrderFormat", m_def_sOrderFormat)
    m_p_sRegNo = PropBag.ReadProperty("p_sRegNo", m_def_p_sRegNo)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("OpenPW", m_OpenPW, m_def_OpenPW)
    Call PropBag.WriteProperty("EditPW", m_EditPW, m_def_EditPW)
    Call PropBag.WriteProperty("EqName", m_EqName, m_def_EqName)
    Call PropBag.WriteProperty("bUseBarcode", m_bUseBarcode, m_def_bUseBarcode)
    Call PropBag.WriteProperty("iPhase", m_iPhase, m_def_iPhase)
    Call PropBag.WriteProperty("iSendPhase", m_iSendPhase, m_def_iSendPhase)
    Call PropBag.WriteProperty("sTestMode", m_sTestMode, m_def_sTestMode)
    Call PropBag.WriteProperty("iFrameN", m_iFrameN, m_def_iFrameN)
    Call PropBag.WriteProperty("p_sID", m_p_sID, m_def_p_sID)
    Call PropBag.WriteProperty("p_sSeq", m_p_sSeq, m_def_p_sSeq)
    Call PropBag.WriteProperty("p_sRack", m_p_sRack, m_def_p_sRack)
    Call PropBag.WriteProperty("p_sPos", m_p_sPos, m_def_p_sPos)
    Call PropBag.WriteProperty("p_iOrdCnt", m_p_iOrdCnt, m_def_p_iOrdCnt)
    Call PropBag.WriteProperty("p_sTIFCd", m_p_sTIFCd, m_def_p_sTIFCd)
    Call PropBag.WriteProperty("sVersion", m_sVersion, m_def_sVersion)
    Call PropBag.WriteProperty("sRstFileNm", m_sRstFileNm, m_def_sRstFileNm)
    Call PropBag.WriteProperty("sRstFilePath", m_sRstFilePath, m_def_sRstFilePath)
    Call PropBag.WriteProperty("Settings", m_Settings, m_def_Settings)

    Call PropBag.WriteProperty("p_sFormula", m_p_sFormula, m_def_p_sFormula)
    Call PropBag.WriteProperty("sOrderFormat", m_sOrderFormat, m_def_sOrderFormat)
    Call PropBag.WriteProperty("p_sRegNo", m_p_sRegNo, m_def_p_sRegNo)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get OpenPW() As String
    OpenPW = m_OpenPW
End Property

Public Property Let OpenPW(ByVal New_OpenPW As String)
    m_OpenPW = New_OpenPW
    PropertyChanged "OpenPW"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EditPW() As String
    EditPW = m_EditPW
End Property

Public Property Let EditPW(ByVal New_EditPW As String)
    m_EditPW = New_EditPW
    PropertyChanged "EditPW"
End Property

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    m_OpenPW = m_def_OpenPW
    m_EditPW = m_def_EditPW
    m_EqName = m_def_EqName
    m_bUseBarcode = m_def_bUseBarcode
    m_iPhase = m_def_iPhase
    m_iSendPhase = m_def_iSendPhase
    m_sTestMode = m_def_sTestMode
    m_iFrameN = m_def_iFrameN
    m_p_sID = m_def_p_sID
    m_p_sSeq = m_def_p_sSeq
    m_p_sRack = m_def_p_sRack
    m_p_sPos = m_def_p_sPos
    m_p_iOrdCnt = m_def_p_iOrdCnt
    m_p_sTIFCd = m_def_p_sTIFCd
    m_sVersion = m_def_sVersion
    m_sRstFileNm = m_def_sRstFileNm
    m_sRstFilePath = m_def_sRstFilePath
    m_Settings = m_def_Settings

    m_p_sFormula = m_def_p_sFormula
    m_sOrderFormat = m_def_sOrderFormat
    m_p_sRegNo = m_def_p_sRegNo
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EqName() As String
    EqName = m_EqName
End Property

Public Property Let EqName(ByVal New_EqName As String)
    m_EqName = New_EqName
    PropertyChanged "EqName"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get bUseBarcode() As Boolean
    bUseBarcode = m_bUseBarcode
End Property

Public Property Let bUseBarcode(ByVal New_bUseBarcode As Boolean)
    m_bUseBarcode = New_bUseBarcode
    PropertyChanged "bUseBarcode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iPhase() As Integer
    iPhase = m_iPhase
End Property

Public Property Let iPhase(ByVal New_iPhase As Integer)
    m_iPhase = New_iPhase
    PropertyChanged "iPhase"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iSendPhase() As Integer
    iSendPhase = m_iSendPhase
End Property

Public Property Let iSendPhase(ByVal New_iSendPhase As Integer)
    m_iSendPhase = New_iSendPhase
    PropertyChanged "iSendPhase"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sTestMode() As String
    sTestMode = m_sTestMode
End Property

Public Property Let sTestMode(ByVal New_sTestMode As String)
    m_sTestMode = New_sTestMode
    PropertyChanged "sTestMode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iFrameN() As Integer
    iFrameN = m_iFrameN
End Property

Public Property Let iFrameN(ByVal New_iFrameN As Integer)
    m_iFrameN = New_iFrameN
    PropertyChanged "iFrameN"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sID() As String
    p_sID = m_p_sID
End Property

Public Property Let p_sID(ByVal New_p_sID As String)
    m_p_sID = New_p_sID
    PropertyChanged "p_sID"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sSeq() As String
    p_sSeq = m_p_sSeq
End Property

Public Property Let p_sSeq(ByVal New_p_sSeq As String)
    m_p_sSeq = New_p_sSeq
    PropertyChanged "p_sSeq"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sRack() As String
    p_sRack = m_p_sRack
End Property

Public Property Let p_sRack(ByVal New_p_sRack As String)
    m_p_sRack = New_p_sRack
    PropertyChanged "p_sRack"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sPos() As String
    p_sPos = m_p_sPos
End Property

Public Property Let p_sPos(ByVal New_p_sPos As String)
    m_p_sPos = New_p_sPos
    PropertyChanged "p_sPos"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get p_iOrdCnt() As Integer
    p_iOrdCnt = m_p_iOrdCnt
End Property

Public Property Let p_iOrdCnt(ByVal New_p_iOrdCnt As Integer)
    m_p_iOrdCnt = New_p_iOrdCnt
    PropertyChanged "p_iOrdCnt"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sTIFCd() As String
    p_sTIFCd = m_p_sTIFCd
End Property

Public Property Let p_sTIFCd(ByVal New_p_sTIFCd As String)
    m_p_sTIFCd = New_p_sTIFCd
    PropertyChanged "p_sTIFCd"
End Property


'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sVersion() As String
    sVersion = m_sVersion
End Property

Public Property Let sVersion(ByVal New_sVersion As String)
    m_sVersion = New_sVersion
    PropertyChanged "sVersion"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sRstFileNm() As String
    sRstFileNm = m_sRstFileNm
End Property

Public Property Let sRstFileNm(ByVal New_sRstFileNm As String)
    m_sRstFileNm = New_sRstFileNm
    PropertyChanged "sRstFileNm"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sRstFilePath() As String
    sRstFilePath = m_sRstFilePath
End Property

Public Property Let sRstFilePath(ByVal New_sRstFilePath As String)
    m_sRstFilePath = New_sRstFilePath
    PropertyChanged "sRstFilePath"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get Settings() As String
Attribute Settings.VB_Description = "전송 속도, 패리티, 데이터 비트, 중단 비트 매개 변수를 반환하거나 설정합니다."
    Settings = m_Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    m_Settings = New_Settings
    PropertyChanged "Settings"
End Property

'
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function RcvRstData(sDBPath$) As Variant

    '--- 사용자 확인
    If m_EditPW <> pEditPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Function
    End If
    '---------------

    If m_EqName = "0" Or m_EqName = "" Then
        RaiseEvent DispMsg("검사장비명을 지정해 주십시오.!!!")
        Exit Function
    End If

    '--- 암호 확인
    If m_OpenPW <> pOpenPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Function
    End If
    '-----------------------

    If Trim(sDBPath) = "" Then
        Exit Function
    End If
    
    Timer1.Enabled = False

    '결과조회
    Call GetResultData_AutoVue(sDBPath)
    
    Timer1.Enabled = True
    
End Function

Private Sub GetResultData_AutoVue(ByVal sRstFilePath As String)
'    On Error GoTo ErrRtn
'
'        Dim FSO As FileSystemObject
'
'        Dim i%
'        Dim objFile As Object
'        Dim sFileNm As String
'        Dim sRtnVal As String
'        Dim sRecType As String
'        Dim sOneRow() As String
'        Dim tmpField()  As String
'        Dim tmpSeqNo$, tmpBarcd$
'        Dim tmpIFCd$, tmpRst$, tmpUnit$
'
'        Set FSO = New FileSystemObject
'
'        If FSO.FolderExists(sRstFilePath) Then
'            For Each objFile In FSO.GetFolder(sRstFilePath).Files
'
'                If UCase(Right(objFile.Name, 3)) = "UPL" Then
'                    Open sRstFilePath & "\" & objFile.Name For Input Shared As #77
'                        sRtnVal = ""
'                        Do While Not EOF(77)
'                            sRtnVal = sRtnVal & Input(1, #77)
'                        Loop
'                    Close #77
'
'                    If sTestMode = "77" Then
'                        RaiseEvent PrintRcvLog(sRtnVal)
'                    End If
'
'                    sOneRow = Split(sRtnVal, vbCrLf)
'
'                    For i = 0 To UBound(sOneRow) - 1
'                        If Trim(sOneRow(i)) = "" Then Exit For
'
'                        sRecType = Left(sOneRow(i), 1)
'
'                        Select Case sRecType
'                            Case "H"
'                            Case "P"
'                            Case "O"
'                                '결과값 등록/화면 표시 처리...
'                                With pResultInfo
'                                    If .RSTCNT > 0 Then
'                                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .OTHER)
'                                    End If
'                                End With
'
'                                tmpBarcd = "": tmpIFCd = "": tmpRst = ""
'                                Call Init_pResultInfo
'
'                                tmpField = Split(sOneRow(i), Chr(124))
'                                tmpBarcd = Trim(tmpField(2))
'                                pResultInfo.ID = tmpBarcd
'
'                            Case "R"
'                                tmpField = Split(sOneRow(i), Chr(124))
'                                tmpIFCd = Trim(tmpField(2))
'                                tmpRst = Trim(tmpField(3))
'
'                                '결과정보 구조체에 저장
'                                With pResultInfo
'                                    '결과값 누적
'                                    .RSTCNT = .RSTCNT + 1
'                                    .IFCD = .IFCD & tmpIFCd & Chr(124)
'                                    .RST1 = .RST1 & tmpRst & Chr(124)
'                                    .RST2 = .RST2 & Chr(124)
'                                    .UNIT = .UNIT & Chr(124)
'                                    .FLAG = .FLAG & Chr(124)
'                                End With
'
'                            Case "M"
'                                tmpField = Split(sOneRow(i), Chr(124))
'                                tmpIFCd = Trim(tmpField(2))
'                                tmpRst = Trim(tmpField(5))
'                                tmpRst = Split(tmpRst, "^")(0)
'
'                                '결과정보 구조체에 저장
'                                With pResultInfo
'                                    '결과값 누적
'                                    .RSTCNT = .RSTCNT + 1
'                                    .IFCD = .IFCD & tmpIFCd & Chr(124)
'                                    .RST1 = .RST1 & tmpRst & Chr(124)
'                                    .RST2 = .RST2 & Chr(124)
'                                    .UNIT = .UNIT & Chr(124)
'                                    .FLAG = .FLAG & Chr(124)
'                                End With
'
'                            Case "L"
'                                '결과값 등록/화면 표시 처리...
'                                With pResultInfo
'                                    If .RSTCNT > 0 Then
'                                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .OTHER)
'                                    End If
'                                End With
'
'                                tmpBarcd = "": tmpIFCd = "": tmpRst = ""
'                                Call Init_pResultInfo
'
'                        End Select
'                    Next i
'
'                    FSO.CopyFile sRstFilePath & "\" & objFile.Name, sRstFilePath & "\processed\" & objFile.Name
'                    FSO.DeleteFile sRstFilePath & "\" & objFile.Name
'                End If
'            Next
'        End If
'
'        Set FSO = Nothing
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("GetResultData_AutoVue 오류발생 - " & Err.Description)
'        Timer1.Enabled = True
'    End If
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sFormula() As String
    p_sFormula = m_p_sFormula
End Property

Public Property Let p_sFormula(ByVal New_p_sFormula As String)
    m_p_sFormula = New_p_sFormula
    PropertyChanged "p_sFormula"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get sOrderFormat() As Variant
    sOrderFormat = m_sOrderFormat
End Property

Public Property Let sOrderFormat(ByVal New_sOrderFormat As Variant)
    m_sOrderFormat = New_sOrderFormat
    PropertyChanged "sOrderFormat"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get p_sRegNo() As Variant
    p_sRegNo = m_p_sRegNo
End Property

Public Property Let p_sRegNo(ByVal New_p_sRegNo As Variant)
    m_p_sRegNo = New_p_sRegNo
    PropertyChanged "p_sRegNo"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function Timer_Start() As Variant
    Timer1.Enabled = True
End Function

