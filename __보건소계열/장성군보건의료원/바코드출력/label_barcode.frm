VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form label 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin MSCommLib.MSComm MSComm1 
      Left            =   990
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RTSEnable       =   -1  'True
      BaudRate        =   38400
   End
End
Attribute VB_Name = "label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dd(10), ee(92) As String
Dim check_10 As Boolean

Private Sub Form_Load()
'###########################################################################
' Date        : 2006-01-16
' Dev         : gabin daddy
' Description : for Posdata ver.2000 AL310 Barcode Printer Module
' Client      : 장성보건의료원 임상병리실
'###########################################################################
    
    Dim strTmp As String
    Dim strBarcode As String
    Dim strName As String
    Dim strCategory As String
    
    dd(10) = Command$
    
    '-- 포스데이터2000버전의 경우..
    'dd(10) = "김길동   01050042OCHEMATOLOGY"
    
    '-- CBC[HEMATOLOGY]만 바코드 출력한다.
    If InStr(dd(10), "HEMATOLOGY") = 0 Then
        End
    Else
        '-- 이렇게 카테고리를 기준으로 항목들의 값을 가져오는 이유는 이름의 자리수가 가변적이기 때문임..
        strCategory = Mid(dd(10), InStr(dd(10), "HEMATOLOGY"))
        strBarcode = Mid(dd(10), InStr(dd(10), "HEMATOLOGY") - 10, 10)
        strName = Trim(Mid(dd(10), 1, 6))
    End If
    
    '---------------------
    '????????용도?????????
    '---------------------
    MSComm1.CommPort = 1
    MSComm1.Settings = "38400,n,8,1"
    MSComm1.PortOpen = True
    MSComm1.PortOpen = False
    
    '----------------------------------------------------------------------
    'Serial Port Open...Must be COM1:
    '----------------------------------------------------------------------
    Open "com1:38400,n,8,1" For Output As #1
    
    '----------------------------------------------------------------------
    ' Serial Port output log write
    '----------------------------------------------------------------------
    Open App.Path + "\" + "barcode.log" For Append As #2
    Print #2, dd(10) & vbNewLine;

    '----------------------------------------------------------------------
    ' BARCODE Label format setting
    ' T : barcode text
    ' B : barcode label
    '----------------------------------------------------------------------
    Print #1, "{D,5}"
    Print #1, "{N,3}"
    Print #1, "{F01,500,240;untitled|"
    Print #1, "T01,I,000,100,220,0101,K,0,1,B,I,0|" '-- 이름
    Print #1, "T02,I,000,300,220,0101,K,0,1,B,I,0|" '-- 항목
    Print #1, "T03,I,000,200,189,0101,7,0,1,B,I,0|" '-- 바코드번호
              '--->T : Text data
                        '------->y,x좌표(start position)
                                '----> 폰트가로,세로크기
                                     '-> 폰트 타입
                                       '-> rotate : 0 or 1
                                         '-> row rotate : 0 ~ 3
                                          '------> 고정값이어도 무방하다.
    Print #1, "B04,I,000,100,150,8,8,1,120,0|"
              '--->B : Barcode data
                        '-------> y,x좌표(start position)
                                '---> code128
                                    '-> 1 : 90 rotate
                                       '--> barcode height
    Print #1, "}"
    
    '----------------------------------------------------------------------
    ' BARCODE 출력
    '----------------------------------------------------------------------
    Print #1, "{B01,1,0,1,1,0,C;untitled|"
    Print #1, "T01;" & strName & "|"
    Print #1, "B04;" & strBarcode & "|"
    Print #1, "T02;" & strCategory & "|"
    Print #1, "T03;" & strBarcode & "|"
    Print #1, "}"

    Close #1
    Close #2

    End

End Sub

