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


Function GetBioradQCVal(ByVal strIntBase As String, ByVal strResult As String, ByVal strLevel As String) As String
    Dim strQCVal    As String
    
    GetBioradQCVal = ""
    strQCVal = ""
    
    Select Case UCase(strIntBase)
        Case "GLU"
            If strLevel = "1" Then
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "TRACE":       strQCVal = "0.1 g/dL or 5.5 mmol/L or Trace or 100 mg/dL"
                    Case "+/-", "+-":   strQCVal = "0.1 g/dL or 5.5 mmol/L or Trace or 100 mg/dL"
                    Case "-/+", "-+":   strQCVal = "0.1 g/dL or 5.5 mmol/L or Trace or 100 mg/dL"
                    Case "POSITIVE":    strQCVal = "0.25 g/dL or 14 mmol/L or 1/4% or 250 mg/dL or +"
                    Case "1+":          strQCVal = "0.25 g/dL or 14 mmol/L or 1/4% or 250 mg/dL or +"
                    Case "2+":          strQCVal = "0.5 g/dL or 28 mmol/L or 1/2% or 500 mg/dL or ++"
                    Case "3+":          strQCVal = "1 g/dL or 55 mmol/L or 1.0% or 1000 mg/dL or +++"
                    Case "4+":          strQCVal = "2 g/dL or 111 mmol/L or 2.0% or 2000 mg/dL or ++++"
                End Select
            Else
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "TRACE":       strQCVal = "0.1 g/dL or 5.5 mmol/L or Trace or 100 mg/dL"
                    Case "+/-", "+-":   strQCVal = "0.1 g/dL or 5.5 mmol/L or Trace or 100 mg/dL"
                    Case "-/+", "-+":   strQCVal = "0.1 g/dL or 5.5 mmol/L or Trace or 100 mg/dL"
                    Case "POSITIVE":    strQCVal = "0.25 g/dL or 14 mmol/L or 1/4% or 250 mg/dL or +"
                    Case "1+":          strQCVal = "0.25 g/dL or 14 mmol/L or 1/4% or 250 mg/dL or +"
                    Case "2+":          strQCVal = "0.5 g/dL or 28 mmol/L or 1/2% or 500 mg/dL or ++"
                    Case "3+":          strQCVal = "1 g/dL or 55 mmol/L or 1.0% or 1000 mg/dL or +++"
                    Case "4+":          strQCVal = "2 g/dL or 111 mmol/L or 2.0% or 2000 mg/dL or ++++"
                End Select
            End If
        Case "BIL"
            If strLevel = "1" Then
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative or -"
                    Case "NEGATIVE":    strQCVal = "Negative or -"
                    Case "POSITIVE":    strQCVal = "Small or +"
                    Case "1+":          strQCVal = "Small or +"
                    Case "2+":          strQCVal = "Moderate or ++"
                    Case "3+":          strQCVal = "Large or +++"
                End Select
            Else
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative or -"
                    Case "NEGATIVE":    strQCVal = "Negative or -"
                    Case "POSITIVE":    strQCVal = "Small or +"
                    Case "1+":          strQCVal = "Small or +"
                    Case "2+":          strQCVal = "Moderate or ++"
                    Case "3+":          strQCVal = "Large or +++"
                End Select
            End If
        Case "KET"
            If strLevel = "1" Then
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative or -"
                    Case "NEGATIVE":    strQCVal = "Negative or -"
                    Case "TRACE":       strQCVal = "Trace or 5 mg/dL or 0.5 mmol/L"
                    Case "+/-", "+-":   strQCVal = "Trace or 5 mg/dL or 0.5 mmol/L"
                    Case "-/+", "-+":   strQCVal = "Trace or 5 mg/dL or 0.5 mmol/L"
                    Case "POSITIVE":    strQCVal = "Small or 15 mg/dL or 1.5 mmol/L or +"
                    Case "1+":          strQCVal = "Small or 15 mg/dL or 1.5 mmol/L or +"
                    Case "2+":          strQCVal = "Moderate or 40 mg/dL or 4 mmol/L or ++"
                    Case "3+":          strQCVal = "Large or 80 mg/dL or 8 mmol/L or +++"
                    Case "4+":          strQCVal = "Large or 160 mg/dL or 16 mmol/L or ++++"
                End Select
            Else
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative or -"
                    Case "NEGATIVE":    strQCVal = "Negative or -"
                    Case "TRACE":       strQCVal = "Trace or 5 mg/dL or 0.5 mmol/L"
                    Case "+/-", "+-":   strQCVal = "Trace or 5 mg/dL or 0.5 mmol/L"
                    Case "-/+", "-+":   strQCVal = "Trace or 5 mg/dL or 0.5 mmol/L"
                    Case "POSITIVE":    strQCVal = "Small or 15 mg/dL or 1.5 mmol/L or +"
                    Case "1+":          strQCVal = "Small or 15 mg/dL or 1.5 mmol/L or +"
                    Case "2+":          strQCVal = "Moderate or 40 mg/dL or 4 mmol/L or ++"
                    Case "3+":          strQCVal = "Large or 80 mg/dL or 8 mmol/L or +++"
                    Case "4+":          strQCVal = "Large or 160 mg/dL or 16 mmol/L or ++++"
                End Select
            End If
        Case "SG"
            GetBioradQCVal = strResult
        
        Case "BLO"
            If strLevel = "1" Then
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative or -"
                    Case "NEGATIVE":    strQCVal = "Negative or -"
                    Case "TRACE":       strQCVal = "Non-Hemolyzed Trace"
                    Case "+/-", "+-":   strQCVal = "Non-Hemolyzed Trace"
                    Case "-/+", "-+":   strQCVal = "Non-Hemolyzed Trace"
                    Case "POSITIVE":    strQCVal = "Small or +"
                    Case "1+":          strQCVal = "Small or +"
                    Case "2+":          strQCVal = "Moderate ++"
                    Case "3+":          strQCVal = "Large or +++"
                End Select
            Else
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative or -"
                    Case "NEGATIVE":    strQCVal = "Negative or -"
                    Case "TRACE":       strQCVal = "Non-Hemolyzed Trace"
                    Case "+/-", "+-":   strQCVal = "Non-Hemolyzed Trace"
                    Case "-/+", "-+":   strQCVal = "Non-Hemolyzed Trace"
                    Case "POSITIVE":    strQCVal = "Small or +"
                    Case "1+":          strQCVal = "Small or +"
                    Case "2+":          strQCVal = "Moderate ++"
                    Case "3+":          strQCVal = "Large or +++"
                End Select
            End If
        Case "PH"
            GetBioradQCVal = strResult
            
        Case "PRO"
            If strLevel = "1" Then
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "TRACE":       strQCVal = "Trace"
                    Case "+/-", "+-":   strQCVal = "Trace"
                    Case "-/+", "-+":   strQCVal = "Trace"
                    Case "POSITIVE":    strQCVal = "30 mg/dL or 0.3 g/L or +"
                    Case "1+":          strQCVal = "30 mg/dL or 0.3 g/L or +"
                    Case "2+":          strQCVal = "100 mg/dL or 1.0 g/L or ++"
                    Case "3+":          strQCVal = "300 mg/dL or 3.0 g/L or +++"
                    Case "4+":          strQCVal = "2000 mg/dL or 20.0 g/L or ++++"
                End Select
            Else
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "TRACE":       strQCVal = "Trace"
                    Case "+/-", "+-":   strQCVal = "Trace"
                    Case "-/+", "-+":   strQCVal = "Trace"
                    Case "POSITIVE":    strQCVal = "30 mg/dL or 0.3 g/L or +"
                    Case "1+":          strQCVal = "30 mg/dL or 0.3 g/L or +"
                    Case "2+":          strQCVal = "100 mg/dL or 1.0 g/L or ++"
                    Case "3+":          strQCVal = "300 mg/dL or 3.0 g/L or +++"
                    Case "4+":          strQCVal = "2000 mg/dL or 20.0 g/L or ++++"
                End Select
            End If
            
        Case "URO"
            If strLevel = "1" Then
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "TRACE":       strQCVal = "Trace"
                    Case "+/-", "+-":   strQCVal = "Trace"
                    Case "-/+", "-+":   strQCVal = "Trace"
                    Case "POSITIVE":    strQCVal = "30 mg/dL or 0.3 g/L or +"
                    Case "1+":          strQCVal = "30 mg/dL or 0.3 g/L or +"
                    Case "2+":          strQCVal = "100 mg/dL or 1.0 g/L or ++"
                    Case "3+":          strQCVal = "300 mg/dL or 3.0 g/L or +++"
                    Case "4+":          strQCVal = "2000 mg/dL or 20.0 g/L or ++++"
                End Select
            Else
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "TRACE":       strQCVal = "Trace"
                    Case "+/-", "+-":   strQCVal = "Trace"
                    Case "-/+", "-+":   strQCVal = "Trace"
                    Case "POSITIVE":    strQCVal = "30 mg/dL or 0.3 g/L or +"
                    Case "1+":          strQCVal = "30 mg/dL or 0.3 g/L or +"
                    Case "2+":          strQCVal = "100 mg/dL or 1.0 g/L or ++"
                    Case "3+":          strQCVal = "300 mg/dL or 3.0 g/L or +++"
                    Case "4+":          strQCVal = "2000 mg/dL or 20.0 g/L or ++++"
                End Select
            End If
        
        Case "NIT"
            If strLevel = "1" Then
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "POSITIVE":    strQCVal = "Positive"
                    Case "1+":          strQCVal = "Positive"
                End Select
            Else
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "POSITIVE":    strQCVal = "Positive"
                    Case "1+":          strQCVal = "Positive"
                End Select
            End If

        Case "LEU"
            If strLevel = "1" Then
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "TRACE":       strQCVal = "Trace"
                    Case "+/-", "+-":   strQCVal = "Trace"
                    Case "-/+", "-+":   strQCVal = "Trace"
                    Case "POSITIVE":    strQCVal = "Small or +"
                    Case "1+":          strQCVal = "Small or +"
                    Case "2+":          strQCVal = "Moderate or ++"
                    Case "3+":          strQCVal = "Large or +++"
                End Select
            Else
                Select Case UCase(strResult)
                    Case "-":           strQCVal = "Negative"
                    Case "NEGATIVE":    strQCVal = "Negative"
                    Case "TRACE":       strQCVal = "Trace"
                    Case "+/-", "+-":   strQCVal = "Trace"
                    Case "-/+", "-+":   strQCVal = "Trace"
                    Case "POSITIVE":    strQCVal = "Small or +"
                    Case "1+":          strQCVal = "Small or +"
                    Case "2+":          strQCVal = "Moderate or ++"
                    Case "3+":          strQCVal = "Large or +++"
                End Select
            End If
    End Select
    
    GetBioradQCVal = strQCVal
    
End Function

Sub SendBioRadQC(ByVal strQCResult As Variant)
    
    
    If strQCResult <> "" Then
        Open gHOSP.QCPATH & Format(Now, "yyyymmdd") & ".txt" For Append As #6
        Print #6, strQCResult;
        Close #6
        strQCResult = ""
    End If
     
End Sub

