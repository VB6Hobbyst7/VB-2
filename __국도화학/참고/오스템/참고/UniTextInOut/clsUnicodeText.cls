VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnicodeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Author: Le Duc Hong         http://www.vovisoft.com
Option Explicit
Private mDOMTextFile As DOMDocument
Private mXMLPath As String
Public Function ReadUnicode(TXMLPath)
   Dim objTextFileRoot As IXMLDOMElement
   Set mDOMTextFile = New DOMDocument
   mXMLPath = TXMLPath
   mDOMTextFile.Load mXMLPath
   'start at the root element of the XML
   Set objTextFileRoot = mDOMTextFile.documentElement
   ReadUnicode = objTextFileRoot.nodeTypedValue
End Function
Public Sub WriteUnicode(OutText, Optional TXMLPath)
   Dim tDOMNode As IXMLDOMElement  ' Temporary Node for DOM
   If IsMissing(TXMLPath) Then
      ' Save the information on the screen by creating a new element and add its children to the DOM object
      mDOMTextFile.documentElement.Text = OutText
      ' Update the XML file
      mDOMTextFile.save mXMLPath
   Else
      Set mDOMTextFile = New DOMDocument
      ' Create a Node called "Text" in DOM
      Set tDOMNode = mDOMTextFile.createElement("Text")
      ' Make it the Root Node
      mDOMTextFile.appendChild tDOMNode
      ' Assign Output Text to Root Node
      mDOMTextFile.documentElement.Text = OutText
      ' Update the XML file
      mDOMTextFile.save TXMLPath
   End If
End Sub
