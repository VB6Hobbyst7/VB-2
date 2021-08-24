Attribute VB_Name = "AJUUNI"
Option Explicit

Public Function spUpdateResult(argSpcDate As String, argSpcNo As String, argSpcSeq As String, argEquipExamCode As String, _
                               argRsltCode As String, argRsltEquip As String, argRsltNum As String, argCPM As String, _
                               argEquipCode As String, argEditID As String, argEditIP As String, argCtrlVal1 As String, _
                               argCtrlVal2 As String, argCtrlVal3 As String, argMSG As String, argEquipNameMsg As String) As String
  

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
    
    Dim sRetData    As String
    On Error GoTo ErrExit:
    
    spUpdateResult = ""
    
'''    CREATE OR REPLACE PROCEDURE NETS1.PC_SL_RESULT_IDU
'''    (
'''    in_spcdate in date,   '접수일자
'''    in_spcno in varchar2, '접수번호
'''    in_spcseq in int,     '접수SEQ
'''    in_eqipexamcode in varchar2, '장비코드
'''    in_rsltcode in out varchar2, '결과코드
'''    in_rslteqip in out varchar2, '장비결과
'''    in_rsltnum in out varchar2,  '결과값
'''    in_cpm in varchar2,          '빈값
'''    in_eqipcode in varchar2,     'GEQUIP
'''    in_editid in varchar2,       '아이디
'''    in_editip in varchar2,       '아이피
'''    in_ctrlval1 in varchar2,
'''    in_ctrlval2 in varchar2,
'''    in_ctrlval3 in varchar2,
'''    in_msg in out varchar2,
'''    out_msg in out varchar2,
'''    in_eqipnamemsg in varchar2 default null
'''    )
       
    
       ' Open connection.
    Set cnn1 = New ADODB.Connection
    ' Modify the following line to reflect a Connection within your environment
    '''   strCnn = cn_Ser
    
    ' Create Parameter Objects to be used later

    cnn1.Open "Provider=MSDAORA.1;" & _
              "User ID=" & gDB_Parm.User & ";" & _
              "Password=" & gDB_Parm.Passwd & ";" & _
              "Data Source=" & gDB_Parm.Server & ";" & _
              "Persist Security Info=False"
                            
    Set cmdExeproc = New ADODB.Command
    
    cmdExeproc.ActiveConnection = cnn1
    
    cmdExeproc.CommandText = "{call NETS1.PC_SL_RESULT_IDU(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Set prm1 = cmdExeproc.CreateParameter("in_spcdate", adDate, adParamInput, 19, argSpcDate)
    cmdExeproc.Parameters.Append prm1
    
    Set prm2 = cmdExeproc.CreateParameter("in_spcno", adVarChar, adParamInput, 5, argSpcNo)
    cmdExeproc.Parameters.Append prm2
    
    Set prm3 = cmdExeproc.CreateParameter("in_spcseq", adInteger, adParamInput, 1, argSpcSeq)
    cmdExeproc.Parameters.Append prm3
    
    Set prm4 = cmdExeproc.CreateParameter("in_eqipexamcode", adVarChar, adParamInput, 10, argEquipExamCode)
    cmdExeproc.Parameters.Append prm4
    
    Set prm5 = cmdExeproc.CreateParameter("in_rsltcode", adVarChar, adParamInput, 20, argRsltCode)
    cmdExeproc.Parameters.Append prm5
    
    Set prm6 = cmdExeproc.CreateParameter("in_rslteqip", adVarChar, adParamInput, 20, argRsltEquip)
    cmdExeproc.Parameters.Append prm6
    
    Set prm7 = cmdExeproc.CreateParameter("in_rsltnum", adVarChar, adParamInput, 20, argRsltNum)
    cmdExeproc.Parameters.Append prm7
    
    Set prm8 = cmdExeproc.CreateParameter("in_cpm", adVarChar, adParamInput, 20, argCPM)
    cmdExeproc.Parameters.Append prm8
    
    Set prm9 = cmdExeproc.CreateParameter("in_eqipcode", adVarChar, adParamInput, 10, argEquipCode)
    cmdExeproc.Parameters.Append prm9
    
    Set prm10 = cmdExeproc.CreateParameter("in_editid", adVarChar, adParamInput, 7, argEditID)
    cmdExeproc.Parameters.Append prm10
    
    Set prm11 = cmdExeproc.CreateParameter("in_editip", adVarChar, adParamInput, 30, argEditIP)
    cmdExeproc.Parameters.Append prm11
    
    Set prm12 = cmdExeproc.CreateParameter("in_ctrlval1", adVarChar, adParamInput, 10, argCtrlVal1)
    cmdExeproc.Parameters.Append prm12
    
    Set prm13 = cmdExeproc.CreateParameter("in_ctrlval2", adVarChar, adParamInput, 10, argCtrlVal2)
    cmdExeproc.Parameters.Append prm13
    
    Set prm14 = cmdExeproc.CreateParameter("in_ctrlval3", adVarChar, adParamInput, 10, argCtrlVal3)
    cmdExeproc.Parameters.Append prm14
    
    Set prm15 = cmdExeproc.CreateParameter("in_msg", adVarChar, adParamInput, 10, argMSG)
    cmdExeproc.Parameters.Append prm15
    
    Set prm16 = cmdExeproc.CreateParameter("out_msg", adVarChar, adParamOutput, 200)
    cmdExeproc.Parameters.Append prm16
    
    Set prm17 = cmdExeproc.CreateParameter("in_eqipnamemsg", adVarChar, adParamInput, 10, argEquipNameMsg)
    cmdExeproc.Parameters.Append prm17
    
    ' Now we have the parameters set - execute the command.
    
    cmdExeproc.Execute
    
    ' Show the results
    sRetData = Trim(cmdExeproc.Parameters(16).Value & " ")
    
    cnn1.Close
    spUpdateResult = sRetData
    Exit Function
   
ErrExit:
    cnn1.Close
    spUpdateResult = ""
   
End Function



Public Function spACPT_IDU(argSpcDate As String, argSpcNo As String, argSpcSeq As String, _
                               argEditID As String, argEditIP As String, argDate As String) As String
  

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
    
    Dim sRetData    As String
    On Error GoTo ErrExit:
    
    spACPT_IDU = ""
'''CREATE OR REPLACE PROCEDURE PC_SL_ACPT_IDU
'''        (
'''        in_spcdate in date,
'''        in_spcno in varchar2,
'''        in_spcseq in int,
'''        in_editid in varchar2,
'''        in_editip in varchar2,
'''        in_edittime in date,
'''        out_erryn in out varchar2,
'''        out_msg in out varchar2
'''        )
''  pc_sl_acpt_idu(in_spcdate => :in_spcdate,
''                 in_spcno => :in_spcno,
''                 in_spcseq => :in_spcseq,
''                 in_editid => :in_editid,
''                 in_editip => :in_editip,
''                 in_edittime => :in_edittime,
''                 out_erryn => :out_erryn,
''                 out_msg => :out_msg);
       
    
       ' Open connection.
    Set cnn1 = New ADODB.Connection
    ' Modify the following line to reflect a Connection within your environment
    '''   strCnn = cn_Ser
    
    ' Create Parameter Objects to be used later

    cnn1.Open "Provider=MSDAORA.1;" & _
              "User ID=" & gDB_Parm.User & ";" & _
              "Password=" & gDB_Parm.Passwd & ";" & _
              "Data Source=" & gDB_Parm.Server & ";" & _
              "Persist Security Info=False"
                            
    Set cmdExeproc = New ADODB.Command
    
    cmdExeproc.ActiveConnection = cnn1
    
    cmdExeproc.CommandText = "{call NETS1.pc_sl_acpt_idu(?,?,?,?,?,?,?,?)}"
    
    Set prm1 = cmdExeproc.CreateParameter("in_spcdate", adDate, adParamInput, 19, argSpcDate)
    cmdExeproc.Parameters.Append prm1
    
    Set prm2 = cmdExeproc.CreateParameter("in_spcno", adVarChar, adParamInput, 5, argSpcNo)
    cmdExeproc.Parameters.Append prm2
    
    Set prm3 = cmdExeproc.CreateParameter("in_spcseq", adInteger, adParamInput, 1, argSpcSeq)
    cmdExeproc.Parameters.Append prm3
    
    Set prm4 = cmdExeproc.CreateParameter("in_editid", adVarChar, adParamInput, 7, argEditID)
    cmdExeproc.Parameters.Append prm4
    
    Set prm5 = cmdExeproc.CreateParameter("in_editip", adVarChar, adParamInput, 30, argEditIP)
    cmdExeproc.Parameters.Append prm5
    
    Set prm6 = cmdExeproc.CreateParameter("in_edittime", adDate, adParamInput, 19, argDate)
    cmdExeproc.Parameters.Append prm6
    

    Set prm7 = cmdExeproc.CreateParameter("out_erryn", adVarChar, adParamOutput, 200)
    cmdExeproc.Parameters.Append prm7
    
    Set prm8 = cmdExeproc.CreateParameter("out_msg", adVarChar, adParamOutput, 200)
    cmdExeproc.Parameters.Append prm8
    ' Now we have the parameters set - execute the command.
    
    cmdExeproc.Execute
    
    ' Show the results
    sRetData = Trim(cmdExeproc.Parameters(6).Value & " ")
    
    cnn1.Close
    spACPT_IDU = sRetData
    Exit Function
   
ErrExit:
    cnn1.Close
    spACPT_IDU = ""
   
End Function

