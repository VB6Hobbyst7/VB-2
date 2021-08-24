VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUniInOut 
   Caption         =   "Hex Display"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptUTF8 
      Caption         =   "UTF-8"
      Height          =   375
      Left            =   8490
      TabIndex        =   10
      Top             =   60
      Width           =   795
   End
   Begin VB.OptionButton OptUTF16 
      Caption         =   "Unicode"
      Height          =   375
      Left            =   7260
      TabIndex        =   9
      Top             =   60
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txtHex 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   4890
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   750
      Width           =   5205
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   9330
      TabIndex        =   2
      Top             =   60
      Width           =   765
   End
   Begin VB.CommandButton CmdDisplayHex 
      Caption         =   "DisplayFileHex"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4755
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   1845
      Left            =   90
      TabIndex        =   8
      Top             =   3600
      Width           =   4755
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "8387;3202"
      MatchEntry      =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hex of File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6630
      TabIndex        =   7
      Top             =   450
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text content"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   450
      Width           =   1515
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   4665
      Left            =   60
      TabIndex        =   4
      Top             =   750
      Width           =   4815
      VariousPropertyBits=   -1400879077
      ScrollBars      =   3
      Size            =   "8493;8229"
      FontName        =   "Tahoma"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmUniInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author: Le Duc Hong         http://www.vovisoft.com
Option Explicit
Dim mDOMVowels As clsUnicodeText
Dim UVowels As String
Private Sub CmdBrowse_Click()
' Launch the dialog to let user select a file name
   With CommonDialog1
      .InitDir = GetLocalDirectory  ' Default folder is where program EXE resides
      .ShowOpen  ' Launch the dialog
      txtFileName.Text = .FileName ' Assign selected filename to txtFileName.text
   End With
End Sub
Private Sub CmdDisplayHex_Click()
Dim strFileName As String
' Display the content of the text file in Hex
   Dim TFileEncoding  As coEncoding
   If txtFileName.Text = "" Then
      ' Ask user to select a file
      CmdBrowse_Click
   End If
   strFileName = txtFileName.Text
   ' Obtain the Hex display string and return to txtHex.text
   txtHex.Text = HexDisplayOfFile(strFileName)
   ' Obtain encoding of the Text file
   TFileEncoding = GetFileEncoding(strFileName)
   ' How we read the file depends on its encoding
   Select Case TFileEncoding
   Case coANSI
    ' Do nothing at the moment
   Case coUnicode
      OptUTF16.Value = True  ' set the UTF-16 Radio button to give feedback to user
      TextBox1.Text = TextOfUTF16(strFileName)
   ' Alternatively, you can also read a UTF-16LE by using FileSystemObject
   '       with the next line of code
   '   TextBox1.Text = ReadTextFile(strFileName)
   Case coUTF8
      OptUTF8.Value = True   ' set the UTF-8 Radio button to give feedback to user
      TextBox1.Text = TextofUTF8(strFileName)
   ' Alternatively, you can also read a UTF-8 by
   ' either
   '    using Class clsUnicodeText, see example in Sub Form_Load
   ' or calling Function HexStringToUTF8 like:
   '   TextBox1.Text = HexStringToUTF8(txtHex.Text)
   End Select
End Sub
Function GetFileEncoding(TFileName) As coEncoding
' Return the encoding of the text file: ANSI, UTF-16LE or UTF-8
   Dim b1, FileNum
   On Error Resume Next  ' Ignore any error
   FileNum = FreeFile ' Obtain a File handle from the OS
   Open TFileName For Binary As #FileNum
   b1 = Input(1, #FileNum)   ' Read the first byte
   If Asc(b1) = &HFF Then
      GetFileEncoding = coUnicode  ' UTF-16LE file
   ElseIf Asc(b1) = &HEF Then
      GetFileEncoding = coUTF8  ' UTF-8 file
   Else
      GetFileEncoding = coANSI  ' plain ANSI file
   End If
   Close #FileNum  ' Close the file
End Function
Function TextOfUTF16(TFileName)
' Read byte by byte in raw binary then return the Unicode string
   Dim Dummy, Ch, MSB, FileNum
   Dim TStr As String
   FileNum = FreeFile ' Obtain a File handle from the OS
   Open TFileName For Binary As #FileNum  ' Open Text file as binary
   ' Skip the first two BOM bytes which are &HFF and &HFE
   Dummy = Input(2, #FileNum)
   ' Read through to the end-of-file
   Do While Not EOF(FileNum)
      ' Each Unicode character takes 2 bytes in UTF-16LE file
      Ch = Input(1, #FileNum)    ' Read first byte
      MSB = Input(1, #FileNum)   ' Read second byte which will be Most Significant Byte
      If MSB <> "" Then
         ' Combine the 2 bytes and convert them to Unicode by using Function chrW
         TStr = TStr & ChrW(Asc(MSB) * 256 + Asc(Ch))
      End If
   Loop
   Close #FileNum ' Close file
   TextOfUTF16 = TStr  ' Return the Unicode string
End Function
Function TextofUTF8(TFileName)
' Read byte by byte in raw binary then return the Unicode string
   Dim i, Dummy, FileNum
   Dim Ch As String
   Dim BArray() As Byte
   FileNum = FreeFile ' Obtain a File handle from the OS
   Open TFileName For Binary As #FileNum  ' Open Text file as binary
   ' Skip the first three BOM bytes which are &HEF, &HBB and &HBF
   Dummy = Input(3, #FileNum)
   ' Read through to the end-of-file
   i = 0
   Do While Not EOF(FileNum)
      ReDim Preserve BArray(i)
      Ch = Input(1, #FileNum)
      If Ch <> "" Then
        BArray(i) = Asc(Ch): i = i + 1
      End If
   Loop
   Close #FileNum ' Close file
   ' Convert the byte stream of UTF-8 to Unicode String
   TextofUTF8 = UTF8ToUniStr(BArray)  ' Return the Unicode string

End Function
Function HexStringToUTF8(TStr) As String
' Use Look-up table to convert the Hex display to Unicode String
   Dim i, Text2, TLen, Item, letter
   TStr = Mid(TStr, 10) & " " ' Append an extra blank space
   ' Replace every Unicode 2 or 3 byte Hex pattern by the corresponding Unicode character
   For i = 1 To ListBox1.ListCount - 1
      Item = ListBox1.List(i)  ' Fetch a line form the Listbox
      ' Replace Mid(Item, 2) by first character of Item, which is the Unicode character
      TStr = Replace(TStr, Mid(Item, 2), Left(Item, 1))
   Next
   ' Now tidy up the string
   TLen = Len(TStr)
   i = 1
   Do While i < TLen
      letter = Mid(TStr, i, 1)
      If InStr(UVowels, letter) = 0 Then
         ' Convert Hex of a normal ANSI character to ANSI character itself, eg:  "41" to "A"
         Text2 = Text2 & Chr(HexToDec(Mid(TStr, i, 2)))
         i = i + 3 ' typically "41 " takes 3 bytes , so move up 3 characters
      Else
         ' get here if encountered an actual Unicode, use it as is
         Text2 = Text2 & letter
         i = i + 1 ' move up 1 character
      End If
   Loop
   HexStringToUTF8 = Text2  ' Return the Unicode String
End Function
Sub UnicodeTextToListBox(ByVal Utext, LV)
   Dim Pos
   LV.Clear
   ' Split up into lines to load the Listbox LV
   Pos = InStr(Utext, vbLf)  ' Locate Line Feed character
   Do While Pos > 0
      ' Pluck a line from the left of the UText string and add it to LV
      LV.AddItem Left(Utext, Pos - 1)
      ' Keep the remaining
      Utext = Mid(Utext, Pos + 1)
      ' Locate the next Line Feed character
      Pos = InStr(Utext, vbLf)
   Loop
   LV.AddItem Trim(Utext) & " "  ' Add the last piece of text to Listbox LV
End Sub
Private Sub CmdSave_Click()
' Save content of TextBox1 in either UTF-16LE format or UTF-8 format
' Prompt for Output file name
   With CommonDialog1
      .InitDir = GetLocalDirectory
      .FileName = txtFileName
      .ShowSave
   End With
   If OptUTF16.Value = True Then
      SaveUTF16 TextBox1.Text  ' Save in UTF-16LE format
   Else
      SaveUTF8 TextBox1.Text   ' Save in UTF-8 format
      '   Alternatively, you can also use the Sub SaveUTF8UsingLookUpTable here like:
      ' SaveUTF8UsingLookUpTable  TextBox1.Text
   End If
End Sub
Sub SaveUTF16(TStr)
' Save given Text string in UTF-16LE format
   Dim i As Long, ab() As Byte
   Dim TLen, FileNum
   ' Work out number of bytes required
   TLen = Len(TStr) * 2
   ReDim ab(TLen + 1)  ' Prepare dimension of Byte array
   ' Place BOM of UTF-16LE in first 2 bytes
   ab(0) = &HFF
   ab(1) = &HFE
   ' Copy Unicode String to Byte array, 1 byte at a time
   For i = 0 To TLen - 1
      CopyMemory ab(i + 2), ByVal StrPtr(TStr) + i, 1
   Next
   ' Delete output file if it exists
   If Dir(CommonDialog1.FileName) <> "" Then
     Kill CommonDialog1.FileName
   End If
   FileNum = FreeFile ' Obtain a File handle from the OS
   Open CommonDialog1.FileName For Binary As #FileNum  ' Open output file in binary
   Put #FileNum, , ab  ' Write byte array to file
   Close #FileNum  ' Close the file
End Sub
Sub SaveUTF8(TStr)
' Save given Text string in UTF-8 format
   Dim a(2) As Byte
   Dim BArray() As Byte
   Dim FileNum
   ' Place BOM of UTF-8 in first 3 bytes
   a(0) = &HEF
   a(1) = &HBB
   a(2) = &HBF
   ' Delete output file if it exists
   If Dir(CommonDialog1.FileName) <> "" Then
     Kill CommonDialog1.FileName
   End If
   FileNum = FreeFile ' Obtain a File handle from the OS
   Open CommonDialog1.FileName For Binary As #FileNum
   Put #FileNum, , a  ' Write BOM bytes
   ' Convert the Unicode string to UTF-8 byte array
   BArray = UniStrToUTF8(TStr)
   Put #FileNum, , BArray  ' Write byte array to file
   Close #FileNum  ' Close the file
End Sub

Sub SaveUTF8UsingLookUpTable(TStr)
' Save given Text string in UTF-8 format by using a look-up table
   Dim i, j, k, letter, Pos, Item, TLen, FileNum
   Dim a() As Byte
   ReDim a(2)
   ' Place BOM of UTF-8 in first 3 bytes
   a(0) = &HEF
   a(1) = &HBB
   a(2) = &HBF
   TLen = Len(TStr) ' Fetch length of input string
   j = 3 ' skip BOM bytes
   ' Iterate through every character in the string
   For i = 1 To TLen
      letter = Mid(TStr, i, 1)  ' Fetch a character
      Pos = InStr(UVowels, letter)  ' Locate the character in Unicode Vowel list
      If Pos > 0 Then
         ' yes - it's a Unicode vowel
         ' Fetch the item corresponding to the Vowel from the Listbox
         Item = ListBox1.List(Pos)
         Item = Mid(Item, 2) ' Discard the first character
         ' Convert the Hex string to 2 or 3 UTF-8 bytes
         For k = 1 To Len(Item) Step 3
            ReDim Preserve a(j)  ' make room for new byte in the byte array
            ' Convert Hex to number and place it in a byte of the array
            a(j) = HexToDec(Mid(Item, k, 2))
            j = j + 1 ' Increment dimension of byte array
         Next k
      Else
         ' Get here if it's a normal ANSI character
         ReDim Preserve a(j)
         a(j) = Asc(letter)
         j = j + 1
      End If
   Next
   ' Delete output file if it exists
   If Dir(CommonDialog1.FileName) <> "" Then
     Kill CommonDialog1.FileName
   End If
   FileNum = FreeFile ' Obtain a File handle from the OS
   Open CommonDialog1.FileName For Binary As #FileNum
   Put #FileNum, , a  ' Write byte array to file
   Close #FileNum  ' Close the file
End Sub

Private Sub Form_Load()
   Dim TStr
   Set mDOMVowels = New clsUnicodeText
   ' Read the Unicode Vowel list
   UVowels = mDOMVowels.ReadUnicode(GetLocalDirectory & "UnicodeVowels.xml")
   ' Read Look-up table for UTF-8 bytes
   TStr = mDOMVowels.ReadUnicode(GetLocalDirectory & "VowelsUTF8Hex.xml")
   ' Break up the input string into items of Listbox
   ' ListBox1 is left Visible for your information. Make it invisible if you like.
   UnicodeTextToListBox TStr, ListBox1
End Sub
