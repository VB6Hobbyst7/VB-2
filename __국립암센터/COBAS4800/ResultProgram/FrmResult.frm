VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmResult 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "C4800 Sender"
   ClientHeight    =   2130
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4275
   Icon            =   "FrmResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4275
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   540
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdDisConnect 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dis Con"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4380
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   270
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.FileListBox fList 
      Height          =   1170
      Left            =   5340
      Pattern         =   "*.xml"
      TabIndex        =   5
      Top             =   270
      Width           =   2265
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5460
      Top             =   2370
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   3030
      Width           =   2685
   End
   Begin MSComDlg.CommonDialog cdFilePath 
      Left            =   4530
      Top             =   2310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTrans 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manual Send"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   270
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   270
      Width           =   3705
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4110
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmResCheck 
      Interval        =   5000
      Left            =   4980
      Top             =   2340
   End
   Begin VB.CheckBox chkTime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[ Auto ]"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6660
      TabIndex        =   2
      Top             =   1530
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "Send Type :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4410
      TabIndex        =   10
      Top             =   1590
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   405
      Left            =   6420
      TabIndex        =   7
      Top             =   2250
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "SANSOFT Call Center 0505-300-1544"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   3675
   End
   Begin VB.Label lblState 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Not Connected"
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
      Height          =   225
      Left            =   2250
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "Connection State : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1875
   End
   Begin VB.Menu mnProgram 
      Caption         =   "Program"
      Visible         =   0   'False
   End
   Begin VB.Menu mnSetup 
      Caption         =   "Setup"
      Visible         =   0   'False
      Begin VB.Menu mnSetPath 
         Caption         =   "FilePath"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private Sub Command1_Click()

  Dim bFind As Boolean

  Dim Buff As String

  

  Open "D:\test.txt" For Binary As #1

  bFind = False

  While Not EOF(1) And Not bFind

       Line Input #1, Buff

       If Left(Buff, 3) = "333" Then

          Mid(Buff, 5, 3) = "333"

          Put #1, Seek(1) - Len(Buff & vbCrLf), Buff

          bFind = True

       End If

  Wend

  Close #1

End Sub

Private Sub chkTime_Click()

    If chkTime.Value = "1" Then
        Call WritePrivateProfileString("CONFIG", "AutoSend", "Y", App.Path & "\Interface.ini")
    Else
        Call WritePrivateProfileString("CONFIG", "AutoSend", "N", App.Path & "\Interface.ini")
    End If
    
End Sub

'Private Sub cmdConnect_Click()
'    GetSetUp
'
'    Winsock1.Connect gIP, gPort
'    Timer1.Enabled = True
'
'End Sub
'
'Private Sub cmdDisConnect_Click()
'    Winsock1.Close
'    Timer1.Enabled = False
'
'    Call Timer1_Timer
'End Sub

Private Sub cmdTrans_Click()
    Dim strFileName As String
    Dim i As Double
    Dim strRes As String
    Dim TextLine
    Dim strXMLName As String
    Dim FilNum
    
    Dim varFile         As Variant
    Dim tmpFile         As Variant
    Dim tmpSpecimen     As String
    Dim strSpecimen()   As String
    Dim strData()       As String
    Dim j               As Integer
    Dim k               As Integer
    
    Dim logPoint        As Double
    Dim blnTest         As Boolean
    
    Dim strType16ct     As String
    Dim strType16Res    As String
    Dim strType18ct     As String
    Dim strType18Res    As String
    Dim strTypeOtherct  As String
    Dim strTypeOtherRes As String
    Dim strEquipResult  As String
    
    Dim strTypeLISRes1  As String
    Dim strTypeLISRes2  As String
    Dim strTypeLISRes3  As String
    Dim strTypeRes1     As String
    Dim strTypeRes2     As String
    Dim strTypeRes3     As String
    
    Dim strFileType     As String
    
    cdFilePath.ShowOpen
    
    If cdFilePath.FileName = "" Then
        Exit Sub
    End If
    
    
    strXMLName = cdFilePath.FileName
    
    strFileName = Replace(strXMLName, "xml", "txt")
    
    FileCopy strXMLName, strFileName

    DoSleep 100
    
    If Trim(lblState.Caption) = "Not Connected" Then Exit Sub
    
    tmpFile = ""
    varFile = ""
    j = 0
    blnTest = False
    
    logPoint = FileLen(strXMLName) - 50000
     
    FilNum = FreeFile
    Open strFileName For Input As FilNum ' 파일을 엽니다.
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #FilNum, TextLine ' 변수로 데이터 행을 읽어들입니다.
        TextLine = Replace(TextLine, "癤?", "<")
        tmpFile = tmpFile & TextLine & vbNewLine
        If strFileType = "" And InStr(1, TextLine, "RequestedResult") > 0 Then
            strFileType = mGetP(TextLine, 4, """")
        End If
        If InStr(1, TextLine, "<Tests>") > 0 Then
        '    Winsock1.SendData "</DomainObjects>"
            Seek FilNum, logPoint
            'Exit Do
        End If
        'If InStr(1, TextLine, "<Barcode>") > 0 Then
        '    Winsock1.SendData "</DomainObjects>"
            'Seek FilNum, logPoint
            'Exit Do
        'End If
    
    Loop
    Close #1 ' 파일을 닫습니다



'        If strBarcode = "18082902160" Then
'            Stop
'        End If
                
    strTypeOtherRes = ""
    strType16Res = ""
    strType18Res = ""
    strTypeOtherct = ""
    strType16ct = ""
    strType18ct = ""

    varFile = Split(tmpFile, vbNewLine)
    For i = 0 To UBound(varFile)
        If InStr(varFile(i), "<TestOrder Id=") > 0 Then
            'Debug.Print varFile(i)
            tmpSpecimen = Mid(varFile(i), InStr(varFile(i), "Specimen=") + 10)
            tmpSpecimen = Mid(tmpSpecimen, 1, Len(tmpSpecimen) - 2)
            j = j + 1
            ReDim Preserve strSpecimen(j) As String
            strSpecimen(j) = tmpSpecimen
            blnTest = True
        End If
        
        If blnTest = True And InStr(varFile(i), "<TestResult Id=") > 0 Then
            ReDim Preserve strData(j) As String
            
            If InStr(strFileType, "HPV") > 0 Then
    '          <StringValue Name="SubTest" Value="HPVHRWGT" />
    '          <StringValue Name="Ct:-1" Value="---" /> other ct
    '          <StringValue Name="Ct:1" Value="---" /> 16ct
    '          <StringValue Name="Ct:3" Value="33.8" /> 18ct
    '          <StringValue Name="Ct:5" Value="25.8" />
    '          <StringValue Name="CtDescription:-1" Value="HR HPV" />
    '          <StringValue Name="CtDescription:0" Value="Other HR HPV" />
    '          <StringValue Name="CtDescription:1" Value="HPV-16" />
    '          <StringValue Name="CtDescription:3" Value="HPV-18" />
    '          <StringValue Name="CtDescription:5" Value="BG" />
    '          <StringValue Name="Result 1" Value="NEG Other HR HPV" /> other res
    '          <StringValue Name="Result 2" Value="NEG HPV16" /> 16res
    '          <StringValue Name="Result 3" Value="POS HPV18" /> 18res
                
                '18073004321,C01,---,negative,33.8,positive,---,negative,NEG Other HR HPV; NEG HPV16; POS HPV18,e21a2a04-9520-11e8-8dd0-0023242e6652,
                '18073004321,C01,---,negative,33.8,positive,---,negative,NEG Other HR HPV; NEG HPV16; POS HPV18,e21a2a04-9520-11e8-8dd0-0023242e6652,
    
                '18073004321,C01,---,NEG HPV16,33.8,POS HPV18,---,NEG Other HR HPV,NEG Other HR HPV; NEG HPV16; POS HPV18,e21a2a04-9520-11e8-8dd0-0023242e6652,
                For k = 1 To 17
                    Debug.Print mGetP(varFile(i + 3 + k), 2, """")
                    Select Case mGetP(varFile(i + 3 + k), 2, """")
                        Case "SubTest"
                        'Case "Ct:-1":               strTypeOtherct = mGetP(varFile(i + 3 + k), 4, """")
                        Case "Ct:0":               strTypeOtherct = mGetP(varFile(i + 3 + k), 4, """")
                        Case "Ct:1":                strType16ct = mGetP(varFile(i + 3 + k), 4, """")
                        Case "Ct:3":                strType18ct = mGetP(varFile(i + 3 + k), 4, """")
                        Case "Ct:5"
                        Case "CtDescription:-1"
                        Case "CtDescription:0"
                        Case "CtDescription:1"
                        Case "CtDescription:3"
                        Case "CtDescription:5"
                        Case "Result 1":            strTypeOtherRes = mGetP(varFile(i + 3 + k), 4, """")
                        Case "Result 2":            strType16Res = mGetP(varFile(i + 3 + k), 4, """")
                        Case "Result 3":            strType18Res = mGetP(varFile(i + 3 + k), 4, """")
                    End Select
                Next
                    
                strEquipResult = mGetP(mGetP(varFile(i + 23), 2, "<"), 2, ">")
                If InStr(strType16Res, "NEG") > 0 Then
                    strType16Res = "negative"
                ElseIf InStr(strType16Res, "POS") > 0 Then
                    strType16Res = "positive"
                End If
                
                If InStr(strType18Res, "NEG") > 0 Then
                    strType18Res = "negative"
                ElseIf InStr(strType18Res, "POS") > 0 Then
                    strType18Res = "positive"
                End If
                
                If InStr(strTypeOtherRes, "NEG") > 0 Then
                    strTypeOtherRes = "negative"
                ElseIf InStr(strTypeOtherRes, "POS") > 0 Then
                    strTypeOtherRes = "positive"
                End If
                
                strData(j) = strType16ct & "," & strType16Res & "," & strType18ct & "," & strType18Res & "," & strTypeOtherct & "," & strTypeOtherRes & "," & strEquipResult & "," & strSpecimen(j) & ","
                
                
                strTypeOtherRes = ""
                strType16Res = ""
                strType18Res = ""
                strTypeOtherct = ""
                strType16ct = ""
                strType18ct = ""
            Else
    '          <StringValue Name="AnalysisFlags" Value="" />
    '          <StringValue Name="Result 1" Value="Mutation Detected" />
    '          <StringValue Name="Result 2" Value="Ex19Del;T790M" />
    '          <StringValue Name="LISCode:Result 1" Value="IEGFRP1,IEGFRP1,99ROC" />
    '          <StringValue Name="Result 3" Value="Ex19Del: 13.54;T790M: 5.00" />
    '          <StringValue Name="LISCode:Result 2" Value="IEGFRP101,IEGFRP101,99ROC;IEGFRP104,IEGFRP104,99ROC" />
    '          <StringValue Name="LISCode:Result 3" Value="IEGFRP109,IEGFRP109,99ROC;IEGFRP112,IEGFRP112,99ROC" />
    '
    '
    '
    '          <Interpretation>Mutation Detected</Interpretation>
                
                For k = 1 To 10
                    Debug.Print mGetP(varFile(i + 3 + k), 2, """")
                    Select Case mGetP(varFile(i + 3 + k), 2, """")
                        Case "SubTest"
                        Case "LISCode:Result 1":    strTypeLISRes1 = "" 'mGetP(varFile(i + 3 + k), 4, """")
                        Case "LISCode:Result 2":    strTypeLISRes2 = "" 'mGetP(varFile(i + 3 + k), 4, """")
                        Case "LISCode:Result 3":    strTypeLISRes3 = "" 'mGetP(varFile(i + 3 + k), 4, """")
                        Case "Result 1":            strTypeRes1 = mGetP(varFile(i + 3 + k), 4, """")
                        Case "Result 2":            strTypeRes2 = mGetP(varFile(i + 3 + k), 4, """")
                        Case "Result 3":            strTypeRes3 = mGetP(varFile(i + 3 + k), 4, """")
                    End Select
                Next
                    
                strEquipResult = mGetP(mGetP(varFile(i + 13), 2, "<"), 2, ">")
                strData(j) = strTypeRes1 & "," & strTypeLISRes1 & "," & strTypeRes2 & "," & strTypeLISRes2 & "," & strTypeRes3 & "," & strTypeLISRes3 & "," & strEquipResult & "," & strSpecimen(j) & ","
            
                strTypeLISRes1 = ""
                strTypeLISRes2 = ""
                strTypeLISRes3 = ""
                strTypeRes1 = ""
                strTypeRes2 = ""
                strTypeRes3 = ""

            End If
        End If
    Next
    
    For j = 1 To UBound(strSpecimen)
        For i = 0 To UBound(varFile)
            If InStr(varFile(i), "<Sample Id=") > 0 And InStr(varFile(i), strSpecimen(j)) > 0 Then
                strData(j) = mGetP(mGetP(varFile(i + 3), 2, "<"), 2, ">") & "," & mGetP(mGetP(varFile(i + 5), 2, "<"), 2, ">") & "," & strData(j)
            End If
            If InStr(varFile(i), "<RocheControl Id=") > 0 And InStr(varFile(i), strSpecimen(j)) > 0 Then
                strData(j) = mGetP(mGetP(varFile(i + 3), 2, "<"), 2, ">") & "," & mGetP(mGetP(varFile(i + 5), 2, "<"), 2, ">") & "," & strData(j)
            End If
            
            
        Next
    Next
    
    
    
    Winsock1.SendData Chr(5)
    DoSleep 10
    
    For i = 1 To UBound(strData)
        Winsock1.SendData strData(i) & vbCr
    Next
    
    If InStr(strFileType, "HPV") > 0 Then
        Winsock1.SendData "HPV_Result_Data" & vbCr
    Else
        Winsock1.SendData "eGFR_Result_Data" & vbCr
    End If
    
    Winsock1.SendData Chr(4)
    
    Kill strFileName
    DoSleep 500
        

    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


Private Sub Form_Load()
    GetSetUp
        
    If gAutoSend = "Y" Then
        chkTime.Value = "1"
    Else
        chkTime.Value = "0"
    End If
    
    Winsock1.Connect gIP, gPort
End Sub

Private Sub mnSetPath_Click()
    frmSetup.Show 1
    
End Sub

Private Sub Timer1_Timer()
    If Winsock1.State = 7 Then
        lblState = "Connected"
        lblState.ForeColor = RGB(0, 0, 255)
        
        cmdConnect.Enabled = False
        cmdDisConnect.Enabled = True
    Else
        lblState = "Not Connected"
        lblState.ForeColor = RGB(255, 0, 0)
        Winsock1.Close
        Winsock1.Connect gIP, gPort
        
        cmdConnect.Enabled = True
        cmdDisConnect.Enabled = False
    End If
End Sub

Private Sub tmResCheck_Timer()
    Dim ii As Integer
    Dim i As Double
    Dim strFileName As String
    Dim strXMLName As String
    Dim strRes As String
    Dim TextLine
    Dim FilNum
    
    Dim varFile         As Variant
    Dim tmpFile         As Variant
    Dim tmpSpecimen     As String
    Dim strSpecimen()   As String
    Dim strData()       As String
    Dim j               As Integer
    Dim k               As Integer
    
    Dim logPoint        As Double
    Dim blnTest         As Boolean
    
    Dim strType16ct     As String
    Dim strType16Res    As String
    Dim strType18ct     As String
    Dim strType18Res    As String
    Dim strTypeOtherct  As String
    Dim strTypeOtherRes As String
    Dim strEquipResult  As String
    
    Dim strTypeLISRes1  As String
    Dim strTypeLISRes2  As String
    Dim strTypeLISRes3  As String
    Dim strTypeRes1     As String
    Dim strTypeRes2     As String
    Dim strTypeRes3     As String
    
    Dim strFileType     As String
    
    On Error Resume Next
    
    If chkTime.Value = 0 Then Exit Sub
    If Trim(lblState.Caption) = "Not Connected" Then Exit Sub
    Label3.Caption = "1"
    fList.Refresh
    fList.Path = gFilePath
    
    strRes = ""
    
    For ii = 1 To fList.ListCount
        fList.ListIndex = ii - 1
        strXMLName = gFilePath & "\" & fList.FileName
        strFileName = Replace(strXMLName, "xml", "txt")
        
        FileCopy strXMLName, strFileName
    
        DoSleep 100
        
        If Trim(lblState.Caption) = "Not Connected" Then Exit Sub
        
        tmpFile = ""
        varFile = ""
        j = 0
        blnTest = False
        
        logPoint = FileLen(strXMLName) - 50000
         
        FilNum = FreeFile
        Open strFileName For Input As FilNum ' 파일을 엽니다.
        Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
            Line Input #FilNum, TextLine ' 변수로 데이터 행을 읽어들입니다.
            TextLine = Replace(TextLine, "癤?", "<")
            tmpFile = tmpFile & TextLine & vbNewLine
            If strFileType = "" And InStr(1, TextLine, "RequestedResult") > 0 Then
                strFileType = mGetP(TextLine, 4, """")
            End If
            If InStr(1, TextLine, "<Tests>") > 0 Then
                Seek FilNum, logPoint
            End If
        Loop
        Close #1 ' 파일을 닫습니다
    
        varFile = Split(tmpFile, vbNewLine)
        For i = 0 To UBound(varFile)
            If InStr(varFile(i), "<TestOrder Id=") > 0 Then
                Debug.Print varFile(i)
                tmpSpecimen = Mid(varFile(i), InStr(varFile(i), "Specimen=") + 10)
                tmpSpecimen = Mid(tmpSpecimen, 1, Len(tmpSpecimen) - 2)
                j = j + 1
                ReDim Preserve strSpecimen(j) As String
                strSpecimen(j) = tmpSpecimen
                blnTest = True
            End If
            
            If blnTest = True And InStr(varFile(i), "<TestResult Id=") > 0 Then
                ReDim Preserve strData(j) As String
                
                If InStr(strFileType, "HPV") > 0 Then
                    For k = 1 To 17
                        Debug.Print mGetP(varFile(i + 3 + k), 2, """")
                        Select Case mGetP(varFile(i + 3 + k), 2, """")
                            Case "SubTest"
                            Case "Ct:-1":               strTypeOtherct = mGetP(varFile(i + 3 + k), 4, """")
                            Case "Ct:1":                strType16ct = mGetP(varFile(i + 3 + k), 4, """")
                            Case "Ct:3":                strType18ct = mGetP(varFile(i + 3 + k), 4, """")
                            Case "Ct:5"
                            Case "CtDescription:-1"
                            Case "CtDescription:0"
                            Case "CtDescription:1"
                            Case "CtDescription:3"
                            Case "CtDescription:5"
                            Case "Result 1":            strTypeOtherRes = mGetP(varFile(i + 3 + k), 4, """")
                            Case "Result 2":            strType16Res = mGetP(varFile(i + 3 + k), 4, """")
                            Case "Result 3":            strType18Res = mGetP(varFile(i + 3 + k), 4, """")
                        End Select
                    Next
                        
                    strEquipResult = mGetP(mGetP(varFile(i + 23), 2, "<"), 2, ">")
                    If InStr(strType16Res, "NEG") > 0 Then
                        strType16Res = "negative"
                    ElseIf InStr(strType16Res, "POS") > 0 Then
                        strType16Res = "positive"
                    End If
                    
                    If InStr(strType18Res, "NEG") > 0 Then
                        strType18Res = "negative"
                    ElseIf InStr(strType18Res, "POS") > 0 Then
                        strType18Res = "positive"
                    End If
                    
                    If InStr(strTypeOtherRes, "NEG") > 0 Then
                        strTypeOtherRes = "negative"
                    ElseIf InStr(strTypeOtherRes, "POS") > 0 Then
                        strTypeOtherRes = "positive"
                    End If
                    
                    strData(j) = strType16ct & "," & strType16Res & "," & strType18ct & "," & strType18Res & "," & strTypeOtherct & "," & strTypeOtherRes & "," & strEquipResult & "," & strSpecimen(j) & ","
                Else
                    For k = 1 To 10
                        Debug.Print mGetP(varFile(i + 3 + k), 2, """")
                        Select Case mGetP(varFile(i + 3 + k), 2, """")
                            Case "SubTest"
                            Case "LISCode:Result 1":    strTypeLISRes1 = "" 'mGetP(varFile(i + 3 + k), 4, """")
                            Case "LISCode:Result 2":    strTypeLISRes2 = "" 'mGetP(varFile(i + 3 + k), 4, """")
                            Case "LISCode:Result 3":    strTypeLISRes3 = "" 'mGetP(varFile(i + 3 + k), 4, """")
                            Case "Result 1":            strTypeRes1 = mGetP(varFile(i + 3 + k), 4, """")
                            Case "Result 2":            strTypeRes2 = mGetP(varFile(i + 3 + k), 4, """")
                            Case "Result 3":            strTypeRes3 = mGetP(varFile(i + 3 + k), 4, """")
                        End Select
                    Next
                        
                    strEquipResult = mGetP(mGetP(varFile(i + 13), 2, "<"), 2, ">")
                    strData(j) = strTypeRes1 & "," & strTypeLISRes1 & "," & strTypeRes2 & "," & strTypeLISRes2 & "," & strTypeRes3 & "," & strTypeLISRes3 & "," & strEquipResult & "," & strSpecimen(j) & ","
                
                End If
            End If
        Next
        
        For j = 1 To UBound(strSpecimen)
            For i = 0 To UBound(varFile)
                If InStr(varFile(i), "<Sample Id=") > 0 And InStr(varFile(i), strSpecimen(j)) > 0 Then
                    strData(j) = mGetP(mGetP(varFile(i + 3), 2, "<"), 2, ">") & "," & mGetP(mGetP(varFile(i + 5), 2, "<"), 2, ">") & "," & strData(j)
                End If
                If InStr(varFile(i), "<RocheControl Id=") > 0 And InStr(varFile(i), strSpecimen(j)) > 0 Then
                    strData(j) = mGetP(mGetP(varFile(i + 3), 2, "<"), 2, ">") & "," & mGetP(mGetP(varFile(i + 5), 2, "<"), 2, ">") & "," & strData(j)
                End If
                
                
            Next
        Next
        
        Winsock1.SendData Chr(5)
        DoSleep 10
        
        For i = 1 To UBound(strData)
            Winsock1.SendData strData(i) & vbCr
        Next
        
        If InStr(strFileType, "HPV") > 0 Then
            Winsock1.SendData "HPV_Result_Data" & vbCr
        Else
            Winsock1.SendData "eGFR_Result_Data" & vbCr
        End If
        
        Winsock1.SendData Chr(4)
        
        Kill strFileName
        Kill strXMLName
        DoSleep 500

'''        Label3.Caption = "0"
'''        FileCopy strXMLName, strFileName
'''        DoSleep 500
'''
'''        Label3.Caption = "1"
'''
'''        Kill strXMLName
'''        Label3.Caption = "2"
'''        Dim TextLine
'''        Winsock1.SendData Chr(5)
'''
'''        Open strFileName For Input As #1 ' 파일을 엽니다.
'''        Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
'''            Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
'''            Winsock1.SendData TextLine
'''            If InStr(1, TextLine, "<Tests>") > 0 Then
'''                Winsock1.SendData "</DomainObjects>"
'''                Exit Do
'''            End If
'''
'''            DoSleep 1
'''
'''            Debug.Print TextLine ' 직접 실행 창에 출력합니다.
'''        Loop
'''        Close #1 ' 파일을 닫습니다
'''
'''
'''        Winsock1.SendData Chr(4)
'''
'''        Kill strFileName
''''''        DoSleep 500
        
    Next
End Sub


'WinSock Control ==============================================================================================================
Public Sub WinSock_Listen(argWinSock As Winsock)
    Dim sWinSockPort As String
    
    
    sWinSockPort = gPort
    
    
    If sWinSockPort = "0" Or IsNumeric(sWinSockPort) = False Then
        Exit Sub
    End If
    
    If argWinSock.State <> sckClosed Then
        argWinSock.Close
    End If
    
    argWinSock.LocalPort = gPort
    argWinSock.Listen
    
'''    If EquipNum = 1 Then
'''        lblConnect1.Caption = "연결 대기중..."
'''    Else
'''        lblConnect2.Caption = "연결 대기중..."
'''    End If
    
End Sub

Private Sub Winsock1_Close()
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    Winsock1.LocalPort = gPort
    Winsock1.Listen
    
    
'''    lblConnect1.Caption = "연결 대기중..."
    
End Sub


Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    
    Winsock1.Accept requestID
'''    lblConnect1.Caption = "연결[" & requestID & "]" & Winsock1.RemoteHostIP
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim sTmp As String
    Dim strSendData
    Dim strResFlag
    Dim sSigFlag As String
    Dim sStemp As String
    Dim strResData As String
    Dim strMsgSplit() As String
    Dim strACK As String
    Dim i As Integer
    Dim strMDateTime As String

    Winsock1.GetData sTmp
    
    
End Sub


Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'''    lblConnect1.Caption = "[Error]" & Number & " : " & Description
End Sub



