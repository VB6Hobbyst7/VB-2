
'------------------------------------------------------------------------------
' 은평구용으로 가로 :4.8 세로 :3.6 임     2001. 4. 2
'------------------------------------------------------------------------------
'
'
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
'
DIM dd$(9)
'----------------------------------------------------------------------
'       Serial Port Open...Must be COM1:
'----------------------------------------------------------------------
OPEN "com1:9600,n,8,1,rs,ds,cd" FOR RANDOM AS #1
''OPEN "lll" FOR OUTPUT AS #1
'----------------------------------------------------------------------
'       String define
'----------------------------------------------------------------------
    PRINT #1, "\\1B@A"
    PRINT #1, "\\1Ba1002800390"
    PRINT #1, "\\1By0280"
    PRINT #1, "\\1Bf10"
    PRINT #1, "\\1BE"
    PRINT #1, "\\1Bds10010001550015100000001"
    PRINT #1, "\\1Bds10020001550045100000001"
    PRINT #1, "\\1Bds10030001550075100000001"
    PRINT #1, "\\1Bds10040001550105100000001"
    PRINT #1, "\\1Bds10050001550135100000001"
    PRINT #1, "\\1Bds10060001550165100000001"
    PRINT #1, "\\1Bds10070001550195100000001"
    PRINT #1, "\\1Bds10080001550225100000001"
'----------------------------------------------------------------------
'       Read 'LEO.TXT" & Send to Label Printer
'----------------------------------------------------------------------
OPEN "c:\lhns95\label\leo.txt" FOR INPUT AS #2
LINE INPUT #2, a$
CLOSE #2
FOR i = 1 TO 8
    dd$(i) = MID$(a$, (i - 1) * 10 + 1, 10)
NEXT i
dd$(9) = MID$(a$, 81, 12)
'
    PRINT #1, "\\1Bdw1001"; dd$(1)
    PRINT #1, "\\1Bdw1002"; dd$(2)
    PRINT #1, "\\1Bdw1003"; dd$(3)
    PRINT #1, "\\1Bdw1004"; dd$(4)
    PRINT #1, "\\1Bdw1005"; dd$(5)
    PRINT #1, "\\1Bdw1006"; dd$(6)
    PRINT #1, "\\1Bdw1007"; dd$(7)
    PRINT #1, "\\1Bdw1008"; dd$(8)

    PRINT #1, "\\1B@f10"
    PRINT #1, "\\1Bbs1001050325025012004002111"

    PRINT #1, "\\1Bbw1001"; dd$(9)

PRINT #1, "\\1BP"


CLOSE

'OPEN "lll" FOR INPUT AS #1
'DO UNTIL EOF(1)
'    LINE INPUT #1, a$
'    PRINT a$; : INPUT aaa
'LOOP
'CLOSE #1



END

