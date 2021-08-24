Attribute VB_Name = "mdOraFunc"
Option Explicit


Public Function TransRes_StoreProc(argBarcode As String, argExamCode As String, argSpcCode As String, argReceDate As String, _
                                   argPID As String, argPName As String, argPIO As String, argPSex As String, argPAge As String, _
                                   argResGubun As String, argNumRes As String, argStrRes As String, argState As String) As String

  
   Dim cnn1             As ADODB.Connection
   Dim cmdExeproc       As ADODB.Command

   Dim prm1             As ADODB.Parameter
   Dim prm2             As ADODB.Parameter
   Dim prm3             As ADODB.Parameter
   Dim prm4             As ADODB.Parameter
   Dim prm5             As ADODB.Parameter
   Dim prm6             As ADODB.Parameter
   Dim prm7             As ADODB.Parameter
   Dim prm8             As ADODB.Parameter
   Dim prm9             As ADODB.Parameter
   Dim prm10            As ADODB.Parameter
   Dim prm11            As ADODB.Parameter
   Dim prm12            As ADODB.Parameter
   Dim prm13            As ADODB.Parameter
   Dim prm14            As ADODB.Parameter
   Dim prm15            As ADODB.Parameter
   Dim prm16            As ADODB.Parameter
   Dim prm17            As ADODB.Parameter
   Dim prm18            As ADODB.Parameter
   Dim prm19            As ADODB.Parameter
   
    On Error GoTo ErrExit:
    
    TransRes_StoreProc = ""
    
'''    PROCEDURE UP_LIS_INTERFACE_001_U(P_BCODE_NO INT         --//바코드번호
'''                            ,   P_ORD_CD        VARCHAR2    --//처방코드
'''                            ,   P_SP_CD         VARCHAR2    --//검체코드
'''                            ,   P_REC_YMD       VARCHAR2    --//접수일
'''                            ,   P_PTNT_NO       INT         --//환자번호
'''                            ,   P_PTNT_NM       VARCHAR2    --//환자명
'''                            ,   P_IO_GB         VARCHAR2    --//입/외 구분 10 : 입원, 20: 외래
'''                            ,   P_SEX           VARCHAR2    --//성별
'''                            ,   P_AGE           VARCHAR2    --//나이
'''                            ,   P_RESULT_TYPE   VARCHAR2    --//결과구분 01:숫자형   02:문자형
'''                            ,   P_RESULT_VAL    NUMBER      --//검사결과(숫자형)
'''                            ,   P_RESULT_NM VARCHAR2    --//검사결과(문자형)
'''                            ,   P_HL_GB         VARCHAR2    --//HIGH/LOW 구분  H: HIGH, L: LOW
'''                            ,   P_DPA_GB        VARCHAR2    --//DELTA/PANIC 구분 D : DELTA, P : PANIC
'''                            ,   P_DESC_VALUE    VARCHAR2    --//참고치
'''                            ,   P_STS_CD        VARCHAR2    --//상태 0 : 접수, 1 : 결과전송
'''                            ,   P_EQP_CD        VARCHAR2    --//장비코드  06001
'''                            ,   P_ENT_EMPL_NO   INT         --//보고자      없으면 0
'''                            ,   P_ENT_IP            VARCHAR2 )  --//보고IP      없으면 NULL
                            
                            
       ' Open connection.
   Set cnn1 = New ADODB.Connection
   ' Modify the following line to reflect a Connection within your environment
'''   strCnn = cn_Ser

   ' Create Parameter Objects to be used later

    cnn1.Open "Provider=MSDAORA.1;" & _
              "User ID=" & gDB_Parm.User & ";" & _
              "Password=" & gDB_Parm.Password & ";" & _
              "Data Source=" & gDB_Parm.Server & ";" & _
              "Persist Security Info=False"



    Set cmdExeproc = New ADODB.Command
    
    
    cmdExeproc.ActiveConnection = cnn1
    
    cmdExeproc.CommandText = "{call MCCSI.UP_LIS_INTERFACE_001_U(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
   
    Set prm1 = cmdExeproc.CreateParameter("P_BCODE_NO", adInteger, adParamInput, , argBarcode)
    cmdExeproc.Parameters.Append prm1
    
    Set prm2 = cmdExeproc.CreateParameter("P_ORD_CD", adVarChar, adParamInput, 8, argExamCode)
    cmdExeproc.Parameters.Append prm2
    
    Set prm3 = cmdExeproc.CreateParameter("P_SP_CD", adVarChar, adParamInput, 2, argSpcCode)
    cmdExeproc.Parameters.Append prm3
    
    Set prm4 = cmdExeproc.CreateParameter("P_REC_YMD", adVarChar, adParamInput, 8, argReceDate)
    cmdExeproc.Parameters.Append prm4
    
    
    Set prm5 = cmdExeproc.CreateParameter("P_PTNT_NO", adInteger, adParamInput, , argPID)
    cmdExeproc.Parameters.Append prm5
    
    
    Set prm6 = cmdExeproc.CreateParameter("P_PTNT_NM", adVarChar, adParamInput, 30, argPName)
    cmdExeproc.Parameters.Append prm6
    
    
    Set prm7 = cmdExeproc.CreateParameter("P_IO_GB", adVarChar, adParamInput, 2, argPIO)
    cmdExeproc.Parameters.Append prm7
    
    Set prm8 = cmdExeproc.CreateParameter("P_SEX", adVarChar, adParamInput, 1, argPSex)
    cmdExeproc.Parameters.Append prm8
    
    Set prm9 = cmdExeproc.CreateParameter("P_AGE", adVarChar, adParamInput, 3, argPAge)
    cmdExeproc.Parameters.Append prm9
    
    Set prm10 = cmdExeproc.CreateParameter("P_RESULT_TYPE", adVarChar, adParamInput, 2, argResGubun)
    cmdExeproc.Parameters.Append prm10
    
    
    Set prm11 = cmdExeproc.CreateParameter("P_RESULT_VAL", adNumeric, adParamInput, , argNumRes)
    cmdExeproc.Parameters.Append prm11
    
    Set prm12 = cmdExeproc.CreateParameter("P_RESULT_NM", adVarChar, adParamInput, 30, argStrRes)
    cmdExeproc.Parameters.Append prm12
    
    Set prm13 = cmdExeproc.CreateParameter("P_HL_GB", adVarChar, adParamInput, 1, "")
    cmdExeproc.Parameters.Append prm13
    
    
    Set prm14 = cmdExeproc.CreateParameter("P_DPA_GB", adVarChar, adParamInput, 1, "")
    cmdExeproc.Parameters.Append prm14
    
    Set prm15 = cmdExeproc.CreateParameter("P_DESC_VALUE", adVarChar, adParamInput, 1, "")
    cmdExeproc.Parameters.Append prm15
    
    Set prm16 = cmdExeproc.CreateParameter("P_STS_CD", adVarChar, adParamInput, 1, argState)
    cmdExeproc.Parameters.Append prm16
    
    Set prm17 = cmdExeproc.CreateParameter("P_EQP_CD", adVarChar, adParamInput, 5, "06001")
    cmdExeproc.Parameters.Append prm17
    
    Set prm18 = cmdExeproc.CreateParameter("P_ENT_EMPL_NO", adInteger, adParamInput, , 0)
    cmdExeproc.Parameters.Append prm18
    
    Set prm19 = cmdExeproc.CreateParameter("P_ENT_IP", adVarChar, adParamInput, 15, "")
    cmdExeproc.Parameters.Append prm19
    
    
'''    Set prm2 = cmdExeproc.CreateParameter("Arg_Lot_NO", adVarChar, adParamInput, 20, argLotNo)
'''    cmdExeproc.Parameters.Append prm2

'''
'''    Set prm7 = cmdExeproc.CreateParameter("Return_Err_IDX", adVarChar, adParamOutput, 15)
'''    cmdExeproc.Parameters.Append prm7
    
    cmdExeproc.Execute
    
    TransRes_StoreProc = ""
    
    
    cnn1.Close
    Exit Function
   
ErrExit:
    
    TransRes_StoreProc = "Connect Error"
    
    MsgBox "Unexpected Error: " & Err.Description
    
    Exit Function
End Function


