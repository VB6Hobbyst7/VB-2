VERSION 5.00
Begin VB.Form frmParser 
   Caption         =   "HTML Parser Test Application"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   Icon            =   "frmParser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6540
      TabIndex        =   19
      Top             =   7020
      Width           =   900
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7530
      TabIndex        =   3
      Top             =   7020
      Width           =   900
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source"
      Height          =   2700
      Left            =   120
      TabIndex        =   6
      Top             =   4095
      Width           =   8295
      Begin VB.TextBox txtSource 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   7
         Top             =   585
         Width           =   8055
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title : "
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   270
         Width           =   7695
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Open"
      Height          =   375
      Left            =   5535
      TabIndex        =   1
      Top             =   7020
      Width           =   900
   End
   Begin VB.Frame fraSetup 
      Caption         =   "Setup"
      Height          =   1305
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      Begin VB.CheckBox chkForms 
         Caption         =   "Form"
         Height          =   285
         Left            =   5670
         TabIndex        =   18
         Top             =   900
         Width           =   780
      End
      Begin VB.CheckBox chkScript 
         Caption         =   "Script"
         Height          =   285
         Left            =   7380
         TabIndex        =   17
         Top             =   900
         Width           =   825
      End
      Begin VB.CheckBox chkImage 
         Caption         =   "Image"
         Height          =   285
         Left            =   6480
         TabIndex        =   16
         Top             =   900
         Width           =   825
      End
      Begin VB.CheckBox chkPlugin 
         Caption         =   "PlugIn"
         Height          =   285
         Left            =   4770
         TabIndex        =   15
         Top             =   900
         Width           =   825
      End
      Begin VB.CheckBox chkImbed 
         Caption         =   "Imbed"
         Height          =   285
         Left            =   3870
         TabIndex        =   14
         Top             =   900
         Width           =   825
      End
      Begin VB.CheckBox chkLink 
         Caption         =   "Link"
         Height          =   285
         Left            =   1260
         TabIndex        =   13
         Top             =   900
         Width           =   690
      End
      Begin VB.CheckBox chkFullLink 
         Caption         =   "Full Link"
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   900
         Width           =   1005
      End
      Begin VB.TextBox txtURL 
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Text            =   "http://www.paran.com"
         Top             =   510
         Width           =   8055
      End
      Begin VB.Label lblURL 
         Caption         =   "Page URL :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraLinks 
      Caption         =   "Object"
      Height          =   2550
      Left            =   120
      TabIndex        =   2
      Top             =   1485
      Width           =   8295
      Begin VB.TextBox txtLink 
         Height          =   2235
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   11
         Top             =   225
         Width           =   8055
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   660
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   5295
      Begin VB.Label lblStatus 
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   195
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Logicom Softwares (http://www.logicom.ca)
Option Explicit

Dim objMSHTML As New MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument

Private Sub cmdExit_Click()
    'If MsgBox("Are you sure?", vbYesNo, "Exiting the application") = vbYes Then
        Unload Me
    'End If
End Sub

Private Sub cmdGo_Click()
    Screen.MousePointer = vbHourglass
    
    txtLink = ""
    txtSource = ""
    cmdGo.Enabled = False
    cmdParse.Enabled = False
    cmdExit.Enabled = False
    
    lblStatus.Caption = "Gettting document via HTTP"
    
    ' This function is only available with Internet Explorer 5
    
    Set objDocument = objMSHTML.createDocumentFromUrl(txtURL.Text, vbNullString)
    
    lblStatus.Caption = "Getting and parsing HTML document"
    
    ' Tricky, to make the function wait for the document to complete, usually
    ' the transfer is asynchronus. Note that this string might be different if
    ' you have another language than english for Internet Explorer on the
    ' machine where the code is executed.
    Dim tm As Date, ovtm As Boolean
    tm = Now
    cmdGo.Enabled = False
    Do While objDocument.readyState <> "complete"
        If Now - tm > 10# / 86400 Then
            ovtm = True
            Exit Do
        End If
        DoEvents
    Loop
    
    lblStatus.Caption = "Document completed"
    
    txtSource.Text = objDocument.documentElement.outerHTML
    
    ' Copying the title of the page to the label
    
    lblTitle.Caption = "Title : " & objDocument.Title
    txtLink.Text = ""
    
    txtLink.SelText = "URL = " & objDocument.url & vbCrLf
    txtLink.SelText = "Title = " & objDocument.Title & vbCrLf
    'txtLink.SelText = "Modified = " & objDocument.fileModifiedDate & vbCrLf
    'txtLink.SelText = "Size = " & objDocument.fileSize & vbCrLf
    
    If ovtm Then
        lblStatus.Caption = "Done, but overtime, so may be not fully parsed."
    Else
        lblStatus.Caption = "Done"
    End If
    
    cmdGo.Enabled = True
    cmdParse.Enabled = True
    cmdExit.Enabled = True
    
    Screen.MousePointer = vbNormal
End Sub

Sub cmdParse_Click()
    Dim objLink As HTMLLinkElement
    Dim obj As IHTMLElement
    Dim emb As HTMLEmbed
    Dim frm As HTMLFormElement
    Dim img As HTMLImg
    'Dim Span As HTMLSpanElement
    Dim s As String
    Dim v As Variant
   ' Dim colObj As New Collection
   ' Dim obj1 As Object
    
    'On Error Resume Next
    
    ' Copying the source to the text box
    Screen.MousePointer = vbHourglass
    cmdGo.Enabled = False
    cmdParse.Enabled = False
    cmdExit.Enabled = False
    
    lblStatus.Caption = "Extracting links"
    
    ' Processing the link collection of the HTMLDocument object
    Dim l As String, p As Long
    
    l = Trim(objDocument.location)
    
    'If InStr(l, "?") Then
        p = InStrRev(l, "/")
        If p Then
            l = Left$(l, p)
        End If
    'End If
    
    
    txtLink.Text = ""
    txtLink.SelText = "URL = " & objDocument.url & vbCrLf
    txtLink.SelText = "Title = " & objDocument.Title & vbCrLf
'    txtLink.SelText = "Modified = " & objDocument.fileModifiedDate & vbCrLf
'    txtLink.SelText = "Size = " & objDocument.fileSize & vbCrLf
    
    If chkFullLink = 1 Or chkLink = 1 Then
        txtLink.SelText = vbCrLf & "<links>" & vbCrLf
        For Each objLink In objDocument.links
            s = Trim(objLink)
            If (chkFullLink = 1) Or (Left(s, Len(l)) <> l) Then
                txtLink.SelText = s & vbCrLf
                lblStatus.Caption = "Extracted " & objLink
                lblStatus.Refresh
                DoEvents
            End If
        Next
        
        txtLink.SelText = vbCrLf & "<OnClick>" & vbCrLf
        For Each obj In objDocument.All 'anchors '.documentElement
            If obj Is Nothing Then
            ElseIf obj.children.length > 1 Then
            ElseIf IsNull(obj.onclick) Then
            ElseIf Len(Trim(obj.onclick)) Then
                v = Split(Trim(obj.onclick), """")
                If UBound(v) > 0 Then
                    s = Trim(v(1))
                    p = InStr(UCase(s), "SRC=")
                    If p Then
                        s = Mid$(s, p + 4)
                    End If
                    If chkFullLink Then
                        txtLink.SelText = "<" & obj.tagName & "> " & s & vbCrLf
                    Else
                        txtLink.SelText = s & vbCrLf
                    End If
                ElseIf chkFullLink Then
                    v = Split(Trim(obj.onclick), "'")
                    If UBound(v) > 0 Then
                        txtLink.SelText = "<" & obj.tagName & "> " & v(1) & vbCrLf
                    Else
                        v = Split(Trim(obj.onclick), Chr(10))
                        If UBound(v) > 0 Then
                            txtLink.SelText = "<" & obj.tagName & "> " & v(UBound(v) - 1) & vbCrLf
                        Else
                            txtLink.SelText = "<" & obj.tagName & "> " & v(0) & vbCrLf
                        End If
                    End If
                End If
                DoEvents
            End If
        Next
    End If
    
    If chkImbed = 1 Then
        txtLink.SelText = vbCrLf & "<embeds>" & vbCrLf
        For Each emb In objDocument.embeds
            txtLink.SelText = emb.src & vbCrLf
            lblStatus.Caption = "Extracted " & emb.src
            lblStatus.Refresh
            DoEvents
        Next
    End If
    
    If chkForms = 1 Then
        txtLink.SelText = vbCrLf & "<Forms>" & vbCrLf
        For Each frm In objDocument.Forms
            txtLink.SelText = frm.Name & vbCrLf
            lblStatus.Caption = "Extracted " & frm.Name
            lblStatus.Refresh
            DoEvents
        Next
    End If
    
    If chkImage = 1 Then
        txtLink.SelText = vbCrLf & "<images>" & vbCrLf
        For Each img In objDocument.images
            txtLink.SelText = img.src & vbTab & vbTab & img.fileModifiedDate & vbCrLf
            lblStatus.Caption = "Extracted " & img.src
            lblStatus.Refresh
            DoEvents
        Next
    End If
    
    If chkPlugin = 1 Then
        txtLink.SelText = vbCrLf & "<plugins>" & vbCrLf
        For Each obj In objDocument.plugins
            txtLink.SelText = obj.src & vbCrLf
            lblStatus.Caption = "Extracted " & obj.src
            lblStatus.Refresh
            DoEvents
        Next
    End If
    
    If chkScript = 1 Then
        txtLink.SelText = vbCrLf & "<scripts>" & vbCrLf
        For Each obj In objDocument.scripts
            s = Trim(obj.src)
            If Len(s) Then
                txtLink.SelText = s & vbCrLf
                lblStatus.Caption = "Extracted " & s
            ElseIf Len(Trim(obj.Title)) Then
                txtLink.SelText = obj.Title & vbCrLf
            ElseIf Len(Trim(obj.className)) Then
                txtLink.SelText = obj.className & vbCrLf
            ElseIf Len(Trim(obj.innerHTML)) Then
                s = Trim(obj.innerHTML)
                v = Split(s, vbLf)
                txtLink.SelText = "Gloval" & vbTab '& vbCrLf
                Dim v1 As Variant
                Dim isFun As Boolean, isCmt As Boolean
                For Each v1 In v
                    If Not isCmt And Left$(v1, 2) = "//" Then
                        s = vbTab & v1 & vbCrLf
                        txtLink.SelText = s
                        isCmt = True
                    ElseIf Not isFun And Left$(v1, 3) = "var" Then
                        s = vbTab & v1 & vbCrLf
                        txtLink.SelText = s
                    ElseIf Left$(v1, 3) = "fun" Then
                        s = vbTab & v1 & vbCrLf
                        txtLink.SelText = s
                        isFun = True
                    ElseIf InStr(v1, vbTab & "function ") Then
                        s = vbTab & v1 & vbCrLf
                        txtLink.SelText = s
                        isFun = True
                    End If
                Next
                
            ElseIf Len(Trim(obj.innerText)) Then
                txtLink.SelText = obj.innerText & vbCrLf
            ElseIf Len(Trim(obj.tagurl)) Then
                txtLink.SelText = obj.tagurl & vbCrLf
            Else
                txtLink.SelText = "Blank" & vbCrLf
            End If
            lblStatus.Refresh
            DoEvents
        Next
    End If
            
    cmdGo.Enabled = True
    cmdParse.Enabled = True
    cmdExit.Enabled = True
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is TextBox Then
            If c.MultiLine = False Then
                c.Text = GetSetting("WebParser", "Controls", c.Name, c.Text)
            End If
        ElseIf TypeOf c Is CheckBox Then
            c.Value = GetSetting("WebParser", "Controls", c.Name, c.Value)
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is TextBox Then
            If c.MultiLine = False Then
                SaveSetting "WebParser", "Controls", c.Name, c.Text
            End If
        ElseIf TypeOf c Is CheckBox Then
            SaveSetting "WebParser", "Controls", c.Name, c.Value
        End If
    Next

End Sub
