Attribute VB_Name = "Mod_Function"
'***********************************************************************************
'***  Function Name : Mod_Function
'***  Description   : URL ¿¬°á Module
'***  Function      : S_HomePage
'***  Modification Log : 2006/03/20  ±èµ¿ÈÄ  Initial Coding
'***********************************************************************************

Public Sub S_HomePage(ByVal as_URL As String)

   Dim loIE As Object
   
   On Error Resume Next
   
   Set loIE = CreateObject("InternetExplorer.Application")
   loIE.Visible = True
   loIE.Navigate as_URL
   Set loIE = Nothing
   
End Sub
