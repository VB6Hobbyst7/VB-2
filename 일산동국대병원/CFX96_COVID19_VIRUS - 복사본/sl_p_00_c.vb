Function sl_p_00_main& (ptno_in$(),orddte_in$(),hspcl_in$(),patsect_in$(),frctcd_in$(),tstcd_in$(),dpcd_in$(),user_in$,inscls_in$(),spcno_out$(),msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_00")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"ptno_in",ptno_in())
		rv% = dce_push_array(Socket,"orddte_in",orddte_in())
		rv% = dce_push_array(Socket,"hspcl_in",hspcl_in())
		rv% = dce_push_array(Socket,"patsect_in",patsect_in())
		rv% = dce_push_array(Socket,"frctcd_in",frctcd_in())
		rv% = dce_push_array(Socket,"tstcd_in",tstcd_in())
		rv% = dce_push_array(Socket,"dpcd_in",dpcd_in())
		rv% = dce_push_string(Socket,"user_in",user_in)
		rv% = dce_push_array(Socket,"inscls_in",inscls_in())
		dce_table = dce_submit("sl_p_00","sl_p_00_main",Socket)
	End If
	sl_p_00_main = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spcno_out",spcno_out())
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


