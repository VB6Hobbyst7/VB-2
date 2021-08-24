VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
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
   End
End
Attribute VB_Name = "label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dd(10), ee(92) As String
Dim check_10 As Boolean


'Format1.dat : Leble의 8개 String에 관련된 Leo Command File
'       예)
'           \\1B@A
'           \\1Ba1002800390
'           \\1By0280
'           \\1Bf10
'           \\1BE
'           \\1Bds10010002850015101200001
'           \\1Bds10020002850045101200001
'           \\1Bds10030002850075101200001
'           \\1Bds10040002850105101200001
'           \\1Bds10050002850135101200001
'           \\1Bds10060002850165101200001
'           \\1Bds10070002850195101200001
'           \\1Bds10080002850225101200001
'                   ||
'                   Field No...반드시 01-08
'
'
'Format2.dat : Leble의 1개 Barcode에 관련된 Leo Command File
'       예)
'           \\1B@f10
'           \\1Bbs1001150395023512004002101
'                   ||
'                   Field No...반드시 01
'
'Leo.txt     : Leble에 출력될 8개 String과 1개의 Barcode Data
'       Format : Str-1(10) Str-2(10) ..... Str-8(10) Barcode(12)
'                Total 92 Bytes
'       예)
'           0010020001홍길동    AIDSS     STD       HBe       RPR       매독      당뇨      123456789012
'
'
'
'순서
'       1. "FORMAT1.DAT" 의 전 Line을 Printer로 Send
'       2. "LEO.TXT" 를 Read하여 Format-No, String, Barcode dat를 구해
'           다음과 같이 Printer로 Send
'               chr$(27) "dw" format-no "01" str-1
'               chr$(27) "dw" format-no "02" str-2
'               chr$(27) "dw" format-no "03" str-3
'               chr$(27) "dw" format-no "04" str-4
'               chr$(27) "dw" format-no "05" str-5
'               chr$(27) "dw" format-no "06" str-6
'               chr$(27) "dw" format-no "07" str-7
'               chr$(27) "dw" format-no "08" str-8
'       3. "FORMAT2.DAT" 의 전 Line을 Printer로 Send
'       4. Barcode data send
'               chr$(27) "bw" format-no "01" barcode-data
'       5. Label 출력 Command Send
'               chr$(27) "P"
'
'
'==========================================================================
'\1B@f01
'     01                        Format No
'\1Ba0102900100
'    01                         Froamt No
'      0290                     Width Label
'          0100                 Length Label
'\1By0290
'\1Bf01
'Strng Define
'   1bds01020202000001162200001" + cr$ + lf$
'     ds                        Define (ds=Char, bs=Barcode)
'       01                      Format No(01)
'         02                    Field No(02)
'           02                  Option
'             0200              Start Pos. Garo
'                 0001          Start Pos. Sero
'                     16        Data Length
'                       22      Enlarge Rate (2 x 2)...Garo x Sero
'                         00    Rotate
'                           00  00-Norml, 10-Reverse
'                             1 Font Type (1=KS Gothic)
'Barcode Define
'   1Bbs01050F028001100500400110000  + cr$ + lf$
'     bs                                Barcode set
'       01                              Format No
'         05                            Field
'           0f                          option
'             0280                      X-Position
'                 0110                  Y-Position
'                     05                Barcode Data Length
'                       0040            Barcode Height
'                           01          Barcode 종류(0=39, 1=93, 7=upc-a...)see manual P39
'                             1         Barcode Narrow Bar Unit(dot)
'                              0        Ratio(0=1:2)
'                               0       Rotate(0=Normal), 1=90 ?)
'                                00     Link Option
'



Private Sub Form_Load()
MSComm1.CommPort = 1
MSComm1.Settings = "9600,n,8,1"
MSComm1.PortOpen = True
dd(10) = Command$
'----------------------------------------------------------------------
'       Serial Port Open...Must be COM1:
'----------------------------------------------------------------------





'MSComm1.PortOpen = False
'
''Open "com1:9600,n,8,1,rs,ds,cd" For Random As #1
'Open "com1:9600,n,8,1,rs,ds,cd" For Output As #1
''----------------------------------------------------------------------
''       String define
''----------------------------------------------------------------------
'Error_1:
'Print #1, "\\1B@A"
'Print #1, "\\1Ba1002800390"
'Print #1, "\\1By0280"
'Print #1, "\\1Bf10"
'Print #1, "\\1BE"
'Print #1, "\\1Bds10010002950020102200001"
'Print #1, "\\1Bds10020002950050102200001"
'Print #1, "\\1Bds10030002950082102200001"
'Print #1, "\\1Bds10040002950112102200001"
'Print #1, "\\1Bds10050002950142102200001"
'Print #1, "\\1Bds10060002950172102200001"
'Print #1, "\\1Bds10070002950202102200001"
'Print #1, "\\1Bds10080002950232102200001"
''Print #1, "\\1Bds10090002850262102200001"
''Print #1, "\\1Bds10100002850292102200001"
''Print #1, "\\1Bds10110002850322102200001"
''Print #1, "\\1Bds10120002850352102200001"
''Print #1, "\\1Bds10130002850382102200001"
''Print #1, "\\1Bds10140002850412102200001"
''Print #1, "\\1Bds10150002850442102200001"
''Print #1, "\\1Bds10160002850472102200001"
''Print #1, "\\1Bds10170002850502102200001"
''Print #1, "\\1Bds10180002850532102200001"
''Print #1, "\\1Bds10190002850562102200001"
''Print #1, "\\1Bds10200002850592102200001"


'----------------------------------------------------------------------
'       Read 'LEO.TXT" & Send to Label Printer
'----------------------------------------------------------------------
'Open App.Path + "\leo.txt" For Input As #2
'Line Input #2, a$
'Close #2
    
For i = 1 To 92
    ee(i) = Mid(dd(10), i, 1)
Next i



k = 0
j = 1
For i = 1 To 80
k = k + 1
If ee(i) = " " Then
k = 11
check_10 = True
GoTo next_i
End If
check_10 = False
If k > 10 Then
If check_10 = False Then
j = j + 1
k = 1
End If
End If
dd(j) = dd(j) + ee(i)

next_i:

Next i
dd(9) = Right(dd(10), 12)

''
'    Print #1, "\\1Bdw1001"; dd(1)
'    Print #1, "\\1Bdw1002"; dd(2)
'    Print #1, "\\1Bdw1003"; dd(3)
'    Print #1, "\\1Bdw1004"; dd(4)
'    Print #1, "\\1Bdw1005"; dd(5)
'    Print #1, "\\1Bdw1006"; dd(6)
'    Print #1, "\\1Bdw1007"; dd(7)
'    Print #1, "\\1Bdw1008"; dd(8)
'
'    Print #1, "\\1B@f10"
'    Print #1, "\\1Bbs1001050325025012004002111"
'
'    Print #1, "\\1Bbw1001"; dd(9)
'
'Print #1, "\\1BP"

    MSComm1.Output = "N" + Chr(10) + Chr(13) + "JF" + Chr(10) + Chr(13) _
                  + "A280,50,0,9,1,1,N," + Chr(34) + dd(1) + Chr(34) + Chr(10) + Chr(13) _
                  + "A420,50,0,3,1,1,N," + Chr(34) + dd(3) + Chr(34) + Chr(10) + Chr(13) _
                  + "A280,80,0,3,1,1,N," + Chr(34) + Format(Mid(dd(2), 1, 4), "2005-00-00") + Chr(34) + Chr(10) + Chr(13) _
                  + "B260,105,0,3,2,3,90,B," + Chr(34) + dd(9) + Chr(34) + Chr(10) + Chr(13) _
                  + "D10" + Chr(10) + Chr(13) + "P1" + Chr(10) + Chr(13) + "E"
                                    

'                  + "A430,190,0,3,1,1,N," + Chr(34) + dd(3) + Chr(34) + Chr(10) + Chr(13) _
'                  + "A480,50,0,9,1,1,N," + Chr(34) + "진료과" + Chr(34) + Chr(10) + Chr(13) _

                 
    MSComm1.Output = Chr(2) + "v" + Chr(13)
    
    MSComm1.PortOpen = False

Close

End

End Sub

