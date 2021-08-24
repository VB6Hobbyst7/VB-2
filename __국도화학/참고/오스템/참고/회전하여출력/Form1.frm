VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   570
      TabIndex        =   1
      Top             =   1650
      Width           =   2385
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   450
      TabIndex        =   0
      Top             =   690
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit

   Private Const LF_FACESIZE = 32

   Private Type LOGFONT
      lfHeight As Long
      lfWidth As Long
      lfEscapement As Long
      lfOrientation As Long
      lfWeight As Long
      lfItalic As Byte
      lfUnderline As Byte
      lfStrikeOut As Byte
      lfCharSet As Byte
      lfOutPrecision As Byte
      lfClipPrecision As Byte
      lfQuality As Byte
      lfPitchAndFamily As Byte
      lfFaceName As String * LF_FACESIZE
   End Type

   Private Type DOCINFO
      cbSize As Long
      lpszDocName As String
      lpszOutput As String
      lpszDatatype As String
      fwType As Long
   End Type

   Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
   "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

   Private Declare Function SelectObject Lib "gdi32" _
   (ByVal hdc As Long, ByVal hObject As Long) As Long

   Private Declare Function DeleteObject Lib "gdi32" _
   (ByVal hObject As Long) As Long

   Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

   Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) _
   As Long

   Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" _
   (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
   ByVal lpString As String, ByVal nCount As Long) As Long ' or Boolean

   Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" _
   (ByVal hdc As Long, lpdi As DOCINFO) As Long

   Private Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) _
   As Long

   Private Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) _
   As Long

   Private Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) _
   As Long

   Const DESIREDFONTSIZE = 12     ' Could use variable, TextBox, etc.

   Private Sub Command1_Click()
   ' Combine API Calls with the Printer object
      Dim OutString As String
      Dim lf As LOGFONT
      Dim result As Long
      Dim hOldfont As Long
      Dim hPrintDc As Long
      Dim hFont As Long

      Printer.Print "Printer Object"
      hPrintDc = Printer.hdc
      OutString = "Hello World"

      lf.lfEscapement = 1800
      lf.lfHeight = (DESIREDFONTSIZE * -20) / Printer.TwipsPerPixelY
      hFont = CreateFontIndirect(lf)
      hOldfont = SelectObject(hPrintDc, hFont)
      result = TextOut(hPrintDc, 1000, 1000, OutString, Len(OutString))
      result = SelectObject(hPrintDc, hOldfont)
      result = DeleteObject(hFont)

      Printer.Print "xyz"
      Printer.EndDoc
   End Sub

   Private Sub Command2_Click()
   ' Print using API calls only
      Dim OutString As String  'String to be rotated
      Dim lf As LOGFONT        'Structure for setting up rotated font
      Dim temp As String       'Temp string var
      Dim result As Long       'Return value for calling API functions
      Dim hOldfont As Long     'Hold old font information
      Dim hPrintDc As Long     'Handle to printer dc
      Dim hFont As Long        'Handle to new Font
      Dim di As DOCINFO        'Structure for Print Document info

      OutString = "Hello World"   'Set string to be rotated

   ' Set rotation in tenths of a degree, i.e., 1800 = 180 degrees
      lf.lfEscapement = 1800
      lf.lfHeight = (DESIREDFONTSIZE * -20) / Printer.TwipsPerPixelY
      hFont = CreateFontIndirect(lf)  'Create the rotated font
      di.cbSize = 20                  ' Size of DOCINFO structure
      di.lpszDocName = "My Document" ' Set name of print job (Optional)

   ' Create a printer device context
      hPrintDc = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)

      result = StartDoc(hPrintDc, di) 'Start a new print document
      result = StartPage(hPrintDc)    'Start a new page

   ' Select our rotated font structure and save previous font info
      hOldfont = SelectObject(hPrintDc, hFont)

   ' Send rotated text to printer, starting at location 1000, 1000
      result = TextOut(hPrintDc, 1000, 1000, OutString, Len(OutString))

   ' Reset font back to original, non-rotated
      result = SelectObject(hPrintDc, hOldfont)

   ' Send non-rotated text to printer at same page location
      result = TextOut(hPrintDc, 1000, 1000, OutString, Len(OutString))

      result = EndPage(hPrintDc)      'End the page
      result = EndDoc(hPrintDc)       'End the print job
      result = DeleteDC(hPrintDc)     'Delete the printer device context
      result = DeleteObject(hFont)    'Delete the font object
   End Sub

   Private Sub Form_Load()
      Command1.Caption = "API with Printer object"
      Command2.Caption = "Pure API"
   End Sub

            

