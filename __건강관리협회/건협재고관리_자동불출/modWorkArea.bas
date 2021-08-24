Attribute VB_Name = "modWorkArea"
Option Explicit

Public gTblTEST As String, gTblSPC As String, gTblEMP As String, gTblEQP As String
Public gFldTESTCD As String, gFldTESTNM
Public gFldSPCCD As String, gFldSPCNM As String, gFldSPCDAY As String
Public gFldEMPID As String, gFldEMPNM As String
Public gFldEQPCD As String, gFldEQPNM

Public Sub gsWorkAreaSet()

    If gWorkArea Then
        gTblTEST = "S2LAB001"
        gTblSPC = "S2LAB032"
        gTblEMP = "S2COM006"
        gTblEQP = "S2LAB006"
        
        gFldTESTCD = "TESTCD"
        gFldTESTNM = "TESTNM"
        
        gFldSPCCD = "SPCCD"
        gFldSPCNM = "SPCNM"
        gFldSPCDAY = "USEDAY"
        
        gFldEMPID = "EMPID"
        gFldEMPNM = "EMPNM"
        
        gFldEQPCD = "EQPCD"
        gFldEQPNM = "EQPNM"
    Else
        gTblTEST = "TWMED_ITEM"
        gTblSPC = "TWMED_SPEC"
        gTblEMP = "TWMED_USER2006"
        gTblEQP = "TWMED_MACHINE"
        
        gFldTESTCD = "ITEMCODE"
        gFldTESTNM = "ITEMHNM"
        
        gFldSPCCD = "SPECCODE"
        gFldSPCNM = "SPECNAME"
        gFldSPCDAY = "'10'"
        
        gFldEMPID = "USERID"
        gFldEMPNM = "USER_NM"
        
        gFldEQPCD = "MACHCODE"
        gFldEQPNM = "MACHNAME"
    End If

End Sub
