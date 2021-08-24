Function sl_p_02_acpt& (spcno_in$,tstdte_in$,hspcl_in$,frct_cd_in$,spccd_in$,ptno_in$,time_in$,vol_in$,userid_in$,micro_in$,remk_in$,acptno_out&,msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_02")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"spcno_in",spcno_in)
		rv% = dce_push_string(Socket,"tstdte_in",tstdte_in)
		rv% = dce_push_string(Socket,"hspcl_in",hspcl_in)
		rv% = dce_push_string(Socket,"frct_cd_in",frct_cd_in)
		rv% = dce_push_string(Socket,"spccd_in",spccd_in)
		rv% = dce_push_string(Socket,"ptno_in",ptno_in)
		rv% = dce_push_string(Socket,"time_in",time_in)
		rv% = dce_push_string(Socket,"vol_in",vol_in)
		rv% = dce_push_string(Socket,"userid_in",userid_in)
		rv% = dce_push_string(Socket,"micro_in",micro_in)
		rv% = dce_push_string(Socket,"remk_in",remk_in)
		dce_table = dce_submit("sl_p_02","sl_p_02_acpt",Socket)
	End If
	sl_p_02_acpt = dce_pop_long(dce_table,"dce_result")
	acptno_out = dce_pop_long(dce_table,"acptno_out")
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_02_cancel& (spcno_in$,hspcl_in$,tstdte_in$,frctcd_in$,acpt_in&,flag_in$,micro_in$,msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_02")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"spcno_in",spcno_in)
		rv% = dce_push_string(Socket,"hspcl_in",hspcl_in)
		rv% = dce_push_string(Socket,"tstdte_in",tstdte_in)
		rv% = dce_push_string(Socket,"frctcd_in",frctcd_in)
		Call dce_push_long(Socket,"acpt_in",acpt_in)
		rv% = dce_push_string(Socket,"flag_in",flag_in)
		rv% = dce_push_string(Socket,"micro_in",micro_in)
		dce_table = dce_submit("sl_p_02","sl_p_02_cancel",Socket)
	End If
	sl_p_02_cancel = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


