VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1530
      TabIndex        =   6
      Top             =   6690
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1530
      TabIndex        =   5
      Top             =   6150
      Width           =   2055
   End
   Begin VB.CommandButton CmdMakeControl 
      Caption         =   "컨트롤만들기"
      Height          =   1035
      Left            =   3720
      TabIndex        =   3
      Top             =   6180
      Width           =   1425
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      Height          =   465
      Left            =   7170
      TabIndex        =   1
      Top             =   6270
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5985
      ScaleWidth      =   12825
      TabIndex        =   0
      Top             =   -30
      Width           =   12855
      Begin Threed.SSPanel pan 
         Height          =   405
         Left            =   7560
         TabIndex        =   4
         Top             =   2730
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "잘나오나?"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   7140
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Label1"
         Height          =   285
         Left            =   8370
         TabIndex        =   2
         Top             =   4170
         Width           =   1605
      End
   End
   Begin VB.Label lblStatic 
      Caption         =   "항목명"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   7500
      TabIndex        =   9
      Top             =   7290
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "항목명"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   7
      Top             =   6180
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' Prntform sample from BlackBeltVB.com
' http://blackbeltvb.com
'
' Written by Matt Hart
' Copyright 1999 by Matt Hart
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code. Do not copy this sample to a collection, such as
' a CD-ROM archive. You may link directly to the original sample
' using "http://blackbeltvb.com/prntform.htm"
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'
' This sample utilizes a better method than "PrintForm" to print the contents of
' a Form. PrintForm sometimes excludes grid and other controls. This method simulates
' pressing the PrintScrn key, which copies the image of either the form or screen to
' the clipboard. Note that I call the .Clear method of the Clipboard first - that's because
' it might already have text or something on it.
'
' When printing, note that I scale the picture's height to proportionally fit the
' printer's resolution. The procedure would need to be adjusted if the Height of the
' Form/Screen was greater than Printer.ScaleHeight, or if the Height was greater than
' the Width and the Width was greater than Printer.ScaleWidth.
'
' Updated with VK_MENU keypresses - note that NT needs these.

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const VK_MENU As Byte = &H12
Const VK_SNAPSHOT As Byte = &H2C
Const KEYEVENTF_KEYUP = &H2
 

Private Type POINTAPI
        X As Long
        Y As Long
End Type
 
Private LMousePos    As POINTAPI     'SSPanel의 X,Y 좌표

Private WithEvents cmdDynamicButton        As VB.CommandButton
Attribute cmdDynamicButton.VB_VarHelpID = -1



Private Type BITMAP
bmType As Long
bmWidth As Long
bmHeight As Long
bmWidthBytes As Long
bmPlanes As Integer
bmBitsPixel As Integer
bmBits As Long
End Type

 

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim PicBits() As Byte, PicInfo As BITMAP
Dim Cnt As Long, BytesPerLine As Long


Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H13D


'WM_PRINT = 0x317
 




Private Sub cmdPrint_Click()



Picture1.AutoRedraw = True

Picture1.CurrentX = Label1.Left
Picture1.CurrentY = Label1.Top
Picture1.Print Label1.Caption

Picture1.CurrentX = pan.Left
Picture1.CurrentY = pan.Top
Picture1.Print pan.Caption


Picture1.CurrentX = Image1.Left
Picture1.CurrentY = Image1.Top
Picture1.Print Image1.Picture

SendMessage Picture1.hwnd, WM_PAINT, Picture1.hDC, 0
'SendMessage Picture1.hwnd, WM_PRINT, Picture1.hDC, PRF_CHILDREN Or PRF_CLIENT Or PRF_OWNED

Printer.PaintPicture Picture1.Image, 0, 0, Picture1.Width, Picture1.Height
Printer.EndDoc



SavePicture Picture1.Image, "C:\TEST.BMP"

' Picture1.AutoRedraw = True

    
End Sub

'Private Sub cmdPrintScreen_Click()
'    Dim lHeight As Long
'    Clipboard.Clear
'    Call keybd_event(VK_SNAPSHOT, 1, 0, 0)
'    DoEvents
'    Call keybd_event(VK_SNAPSHOT, 1, KEYEVENTF_KEYUP, 0)
'    Printer.Print
'    lHeight = (Printer.ScaleWidth / Screen.Width) * Screen.Height
'    Printer.PaintPicture Clipboard.GetData, 0, 0, Printer.ScaleWidth, lHeight
'    Printer.EndDoc
'    ' SavePicture Clipboard.GetData, Me.Name & ".BMP"
'End Sub





'Private Sub cmdDynamicButton_Click()
'
'    MsgBox "Click"
'
'End Sub

Private Sub CmdMakeControl_Click()

    Dim obj     As Object

'    Set obj = Me.Controls.Add("VB.CommandButton", "Label44", Picture1)
    Set obj = Me.Controls.Add("VB.Label", "lblStatic", Picture1)
    obj.Move 300, 300, 1200, 450
    obj.Caption = Text2.Text
    obj.BackStyle = 0
    obj.Visible = True
    
'    Set cmdDynamicButton = obj
    
    Set obj = Nothing


End Sub



'Private Sub Command1_Click()
''Get information (such as height and width) about the picturebox
'GetObject Picture1.Image, Len(PicInfo), PicInfo
''reallocate storage space
'BytesPerLine = (PicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
'ReDim PicBits(1 To BytesPerLine * PicInfo.bmHeight * 3) As Byte
''Copy the bitmapbits to the array
'GetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
''Invert the bits
'For Cnt = 1 To UBound(PicBits)
'PicBits(Cnt) = 255 - PicBits(Cnt)
'Next Cnt
''Set the bits back to the picture
'SetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
''refresh
'Picture1.Refresh
'End Sub

Private Sub lblMouseDown(X As Single, Y As Single)
    LMousePos.X = X
    LMousePos.Y = Y
End Sub


Private Sub lblMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LPanPos As POINTAPI
     
    If Button = vbLeftButton Or Button = vbRightButton Then
       LPanPos.X = Label1.Left + X - LMousePos.X
       LPanPos.Y = Label1.Top + Y - LMousePos.Y

       LPanPos.X = IIf(LPanPos.X < 0, 0, LPanPos.X)
       LPanPos.Y = IIf(LPanPos.Y < 0, 0, LPanPos.Y)

       Label1.Move LPanPos.X, LPanPos.Y

    End If
 
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    LMousePos.X = X
    LMousePos.Y = Y
    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim LPanPos As POINTAPI

    If Button = vbLeftButton Or Button = vbRightButton Then
       LPanPos.X = Label1.Left + X - LMousePos.X
       LPanPos.Y = Label1.Top + Y - LMousePos.Y

       LPanPos.X = IIf(LPanPos.X < 0, 0, LPanPos.X)
       LPanPos.Y = IIf(LPanPos.Y < 0, 0, LPanPos.Y)
       
       Label1.Move LPanPos.X, LPanPos.Y
    End If
 
End Sub


Private Sub lblStatic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LMousePos.X = X
    LMousePos.Y = Y
End Sub

'Private LMouseDown As Boolean        '마우스 Down 상태의 Flag 저장

 

'{마우스가 눌러진 상태의 이벤트 처리}

Private Sub Pan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    LMousePos.X = X
    LMousePos.Y = Y

End Sub

 

'{마우스가 움직일때  이벤트 처리}

Private Sub Pan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 

    Dim LPanPos As POINTAPI

        

'    If LMouseDown Then

    If Button = vbLeftButton Or Button = vbRightButton Then

    

       LPanPos.X = pan.Left + X - LMousePos.X

       LPanPos.Y = pan.Top + Y - LMousePos.Y

       

       LPanPos.X = IIf(LPanPos.X < 0, 0, LPanPos.X)

       LPanPos.Y = IIf(LPanPos.Y < 0, 0, LPanPos.Y)

       

       pan.Move LPanPos.X, LPanPos.Y

       

    End If

 

End Sub

 

'{마우스가 눌러진 상태가 복귀할때  이벤트 처리}

'Private Sub Pan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'

'    LMouseDown = False

'

'End Sub


