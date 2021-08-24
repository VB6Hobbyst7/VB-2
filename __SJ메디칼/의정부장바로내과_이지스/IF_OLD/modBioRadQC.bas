Attribute VB_Name = "modBioRadQC"
Option Explicit

'Point|201106091200|1|1|418761|33790|289|22|0900|0006|93|6|sa|||5.8|
'Point|201106091200|1|2|418761|33790|289|22|0900|0006|93|6|sa|||9.7|

'              strQCVal = "Point" & "|"
'strQCVal = strQCVal & strDate & "|" '-- Date Time
'strQCVal = strQCVal & "1" & "|"                         ' run
'strQCVal = strQCVal & Mid(strBarno, 6, 1) & "|"         ' level
'strQCVal = strQCVal & "447834" & "|"                    ' lab
'strQCVal = strQCVal & txtLot.Text & "|"                 ' lot
'strQCVal = strQCVal & strAnalyte & "|"                  ' analyte
'strQCVal = strQCVal & "619" & "|"                       ' method // Electrochemiluminescence (ECL)
'strQCVal = strQCVal & "1039" & "|"                      ' instrument // Roche Elecsys
'strQCVal = strQCVal & "0006" & "|"                      ' reagent // Dedicated Reagent
'strQCVal = strQCVal & strUnit & "|"                     ' unit
'strQCVal = strQCVal & "6" & "|"                         ' temperature ==> 6   No Temperature
'strQCVal = strQCVal & "sa" & "|"
'strQCVal = strQCVal & "" & "|"
'strQCVal = strQCVal & "" & "|"
'strQCVal = strQCVal & strResult & "|"
'strQCVal = strQCVal & vbCrLf


'#########################################################################

' QC 별  설정     : lot,
' 장비별 설정     : lab, method, instrument, reagent
' 검사항목별 설정 : analyte, unit
'
'#########################################################################
Function MakeBioRadQC(ByVal pEqpCD As String, _
                      ByVal pRun As String, _
                      ByVal pLevel As String, _
                      ByVal pLab As String, _
                      ByVal pLot As String, _
                      ByVal pAnalyte As String, _
                      ByVal pMethod As String, _
                      ByVal pInstrument As String, _
                      ByVal pReagent As String, _
                      ByVal pUnit As String, _
                      ByVal pTemperature As String, _
                      ByVal pResult As String)

    Dim strQCVal    As String
    Dim strDtTM     As String

    strQCVal = ""
    strDtTM = Format(Now, "yyyymmddhhmm")
    
               strQCVal = "Point" & "|"
    strQCVal = strQCVal & strDtTM & "|"         ' Date Time     // yyyymmddhhmm
    strQCVal = strQCVal & pRun & "|"            ' run           // 1,2,3,4
    strQCVal = strQCVal & pLevel & "|"          ' level         // 1,2,3
    strQCVal = strQCVal & pLab & "|"            ' lab           // 447834(병원코드로 대체 가능?)
    strQCVal = strQCVal & pLot & "|"            ' lot           // 159792(입력)
    strQCVal = strQCVal & pAnalyte & "|"        ' analyte       // 검사항목마다 세팅,  Cyfra 21-1 : pAnalyte = "222"
    strQCVal = strQCVal & pMethod & "|"         ' method        // 619 Electrochemiluminescence (ECL)
    strQCVal = strQCVal & pInstrument & "|"     ' instrument    // Roche Elecsys
    strQCVal = strQCVal & pReagent & "|"        ' reagent       // 0006 : Dedicated Reagent
    strQCVal = strQCVal & pUnit & "|"           ' unit          // 검사항목마다 세팅,  Cyfra 21-1 : pUnit = "2" (ng/mL)
    strQCVal = strQCVal & pTemperature & "|"    ' temperature   // 6 : No Temperature
    strQCVal = strQCVal & "sa" & "|"
    strQCVal = strQCVal & "" & "|"
    strQCVal = strQCVal & "" & "|"
    strQCVal = strQCVal & pResult & "|"
    strQCVal = strQCVal & vbCrLf

    Call SendBioRadQC(strQCVal)

End Function


Sub SendBioRadQC(ByVal strQCResult As Variant)
    
    
    If strQCResult <> "" Then
        Open gHOSP.QCPATH & Format(Now, "yyyymmdd") & ".DAT" For Append As #6
        Print #6, strQCResult;
        Close #6
        strQCResult = ""
    End If
     
End Sub

