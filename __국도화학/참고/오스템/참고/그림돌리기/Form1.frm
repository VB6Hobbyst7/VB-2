VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.PictureBox Picture2 
      Height          =   6615
      Left            =   480
      ScaleHeight     =   6555
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   840
      Width           =   7935
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   8520
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlgBlt Lib "GDI32.dll" (ByVal hDCDest As Long, _
    ByRef lpPoint As PointAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, _
    ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

Private Declare Function CreateCompatibleDC Lib "GDI32.dll" (ByVal hDC As Long) As Long

Private Declare Function SelectObject Lib "GDI32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "GDI32.dll" (ByVal hDC As Long) As Long

Private Type PointAPI
  X As Long
  Y As Long
End Type

Dim xPos As Integer, yPos As Integer, angle As Single, xStep As Integer, yStep As Integer

Private Function DrawStdPictureRot(ByVal inDC As Long, ByVal inX As Long, _
  ByVal inY As Long, ByVal inAngle As Single, ByRef inPicture As StdPicture) As Long
  Dim hDC As Long
  Dim hOldBMP As Long
  Dim PlgPts(0 To 4) As PointAPI
  Dim PicWidth As Long, PicHeight As Long
  Dim HalfWidth As Single, HalfHeight As Single
  Dim AngleRad As Single

  Const Pi As Single = 3.14159
  Const HalfPi As Single = Pi * 0.5

  ' Validate input picture
  If (inPicture Is Nothing) Then Exit Function
  If (inPicture.Type <> vbPicTypeBitmap) Then Exit Function

  ' Get picture size
  PicWidth = ScaleX(inPicture.Width, vbHimetric, vbPixels)
  PicHeight = ScaleY(inPicture.Height, vbHimetric, vbPixels)

  ' Get half picture size and angle in radians
  HalfWidth = PicWidth / 2
  HalfHeight = PicHeight / 2
  AngleRad = (inAngle / 180) * Pi

  ' Create temporary DC and select input picture into it
  hDC = CreateCompatibleDC(0&)
  hOldBMP = SelectObject(hDC, inPicture.Handle)

  If (hOldBMP) Then    ' Get angle vectors for width and height
    PlgPts(0).X = Cos(AngleRad) * HalfWidth
    PlgPts(0).Y = Sin(AngleRad) * HalfWidth
    PlgPts(1).X = Cos(AngleRad + HalfPi) * HalfHeight
    PlgPts(1).Y = Sin(AngleRad + HalfPi) * HalfHeight

    ' Project parallelogram points for rotated area
    PlgPts(2).X = HalfWidth + inX - PlgPts(0).X - PlgPts(1).X
    PlgPts(2).Y = HalfHeight + inY - PlgPts(0).Y - PlgPts(1).Y
    PlgPts(3).X = HalfWidth + inX - PlgPts(1).X + PlgPts(0).X
    PlgPts(3).Y = HalfHeight + inY - PlgPts(1).Y + PlgPts(0).Y
    PlgPts(4).X = HalfWidth + inX - PlgPts(0).X + PlgPts(1).X
    PlgPts(4).Y = HalfHeight + inY - PlgPts(0).Y + PlgPts(1).Y

    ' Draw rotated image
    DrawStdPictureRot = PlgBlt(inDC, PlgPts(2), _
        hDC, 0, 0, PicWidth, PicHeight, 0&, 0, 0)

    ' De-select Bitmap from DC
    Call SelectObject(hDC, hOldBMP)
  End If

  ' Destroy temporary DC
  Call DeleteDC(hDC)
End Function

Private Sub ReDraw()
  Call Picture2.Cls
  Call DrawStdPictureRot(Picture2.hDC, xPos, yPos, angle, Picture1.Picture)
  Call Picture2.Refresh
End Sub

Private Sub Form_Load()
  Picture2.AutoRedraw = True
  Picture2.ScaleMode = vbPixels

  Timer1.Enabled = True
  Timer1.Interval = 25
  xStep = 5
  yStep = 5
End Sub

Private Sub Timer1_Timer()
  If xPos > (Picture2.Width - ScaleX(Picture1.Picture.Width)) / 15 Then
    xStep = -5
  ElseIf xPos < 0 Then
    xStep = 5
  End If

  If yPos > (Picture2.Height - ScaleY(Picture1.Picture.Height)) / 15 Then
    yStep = -5
  ElseIf yPos < 0 Then
    yStep = 5
  End If

  If angle >= 360 Then
    angle = 0
  Else
    angle = angle + 5
  End If

  xPos = xPos + xStep
  yPos = yPos + yStep

  Call ReDraw
End Sub

