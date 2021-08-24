Attribute VB_Name = "modBarcode"
Option Explicit

Public Enum enFontSize
    FONT_DF = "0"                   '
    FONT_SMALL = "1"                '
    FONT_MIDDLE = "2"               '
    FONT_LARGE = "3"                '
End Enum

Public Enum enFontKor
    FONT_BATANG = "0"               '
    FONT_GULIM = "1"                '
End Enum

Public Enum enRotation
    ROT_ZERO = "00"                 '
    ROT_90 = "01"                   '
    ROT_180 = "02"                  '
    ROT_270 = "03"                  '
End Enum

Public Enum enReverse
    SET_REVERSE = "1"               '
    SET_NORMAL = "0"                '
End Enum

Public Enum enBold
    SET_BOLD = "1"                  '
    SET_NORMAL = "0"                '
End Enum

Public Enum enBarStyle
    BAR_CODE39 = "00"               '
    BAR_CODE39C = "01"              '
    BAR_CODE2OF5 = "02"             '
End Enum

Public Enum enBarData
    '일반검체라벨
    DAT_LOCATION = 1                '
    DAT_WORKAREA = 2                '
    DAT_ACCDT = 3                   '
    DAT_ACCSEQ = 4                  '
    DAT_SPCNO = 5                   '
    DAT_DEPT = 6                    '
    DAT_ORDDT = 7                   '
    DAT_COLTM = 8                   '
    DAT_PTNM = 9                    '
    DAT_PTID = 10                   '
    DAT_SPCNM = 11                  '
    DAT_ORDNM1 = 12                 '
    DAT_ORDNM2 = 13                 '
    DAT_FROZEN = 14                 '
End Enum

Public Enum enLabelData
    '혈액백라벨
    DAT_BLDPTNM = 1                 '
    DAT_BLDPTID = 2                 '
    DAT_WARDID = 3                  '
    DAT_COMPNM = 4                  '
    DAT_VOLUME = 5                  '
    DAT_BLDTYPE = 6                 '
    DAT_BLDNO = 7                   '
    DAT_TESTDT = 8                  '
End Enum

Type tpBarData
    Key         As enBarData        '
    ElementNo   As String           '
    PosY        As String           '
    PosX        As String           '
    Length      As String           '
    FontY       As enFontSize       '
    FontX       As enFontSize       '
    BoldFg      As enBold           '
    ReverseFg   As enReverse        '
    PrtFg       As String           '
End Type

Type tpBarcode
    PosY    As String               '
    PosX    As String               '
    Length  As String               '
    Height  As String               '
    Style   As enBarStyle           '
    PrtFg   As String               '
End Type

Type tpStatInfo
    PrtLineFg       As String       '
    PrtReverseFg    As String       '
    ErDeptCd        As String       '
    PosY            As String       '
    PosX            As String       '
    Length          As String       '
    Width           As String       '
    ReverseFld      As enBarData    '
End Type

Public Const BAR_PORT = "BAR01"         ' Port
Public Const BAR_WIDTH = "BAR02"        ' Width
Public Const BAR_LENGTH = "BAR03"       ' Length
Public Const BAR_TOTLEN = "BAR04"       ' Total Length
Public Const BAR_BARCODE = "BAR05"      ' Barcode
Public Const BAR_ACCFG = "BAR06"        ' Accfg
Public Const BAR_STAT = "BAR07"         ' Statfg
Public Const BAR_KIND = "BAR08"         ' Barcode Kine
Public Const COM2_BAR_CONFIG = "Z001"   ' COM002 의 Cdindex
'Global dbconn    As Connection         '
Public objTables    As Object           '
Public objFields    As Object           '
'Public T_COM004     As String           '
'Public T_COM002     As String
