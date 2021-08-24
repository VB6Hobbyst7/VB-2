Function sl_p_01_spcnum& (ptno_in$(),orddte_in$(),hspcl_in$(),frctcd_in$(),frctcd_extra_in$(),spccd_in$(),wkgrp_in$(),ordsite_in$(),dpcd_in$(),rsvdte_in$(),iogb_in$(),remk_in$(),job_in$(),user_in$,place_in$,count_in&,spcno$(),msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"ptno_in",ptno_in())
		rv% = dce_push_array(Socket,"orddte_in",orddte_in())
		rv% = dce_push_array(Socket,"hspcl_in",hspcl_in())
		rv% = dce_push_array(Socket,"frctcd_in",frctcd_in())
		rv% = dce_push_array(Socket,"frctcd_extra_in",frctcd_extra_in())
		rv% = dce_push_array(Socket,"spccd_in",spccd_in())
		rv% = dce_push_array(Socket,"wkgrp_in",wkgrp_in())
		rv% = dce_push_array(Socket,"ordsite_in",ordsite_in())
		rv% = dce_push_array(Socket,"dpcd_in",dpcd_in())
		rv% = dce_push_array(Socket,"rsvdte_in",rsvdte_in())
		rv% = dce_push_array(Socket,"iogb_in",iogb_in())
		rv% = dce_push_array(Socket,"remk_in",remk_in())
		rv% = dce_push_array(Socket,"job_in",job_in())
		rv% = dce_push_string(Socket,"user_in",user_in)
		rv% = dce_push_string(Socket,"place_in",place_in)
		Call dce_push_long(Socket,"count_in",count_in)
		dce_table = dce_submit("sl_p_01","sl_p_01_spcnum",Socket)
	End If
	sl_p_01_spcnum = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spcno",spcno())
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_01_spcnum_sl& (ptno_in$(),orddte_in$(),hspcl_in$(),frctcd_in$(),frctcd_extra_in$(),spccd_in$(),wkgrp_in$(),ordsite_in$(),dpcd_in$(),rsvdte_in$(),iogb_in$(),remk_in$(),job_in$(),day_yn_in$(),user_in$,place_in$,count_in&,spcno$(),msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"ptno_in",ptno_in())
		rv% = dce_push_array(Socket,"orddte_in",orddte_in())
		rv% = dce_push_array(Socket,"hspcl_in",hspcl_in())
		rv% = dce_push_array(Socket,"frctcd_in",frctcd_in())
		rv% = dce_push_array(Socket,"frctcd_extra_in",frctcd_extra_in())
		rv% = dce_push_array(Socket,"spccd_in",spccd_in())
		rv% = dce_push_array(Socket,"wkgrp_in",wkgrp_in())
		rv% = dce_push_array(Socket,"ordsite_in",ordsite_in())
		rv% = dce_push_array(Socket,"dpcd_in",dpcd_in())
		rv% = dce_push_array(Socket,"rsvdte_in",rsvdte_in())
		rv% = dce_push_array(Socket,"iogb_in",iogb_in())
		rv% = dce_push_array(Socket,"remk_in",remk_in())
		rv% = dce_push_array(Socket,"job_in",job_in())
		rv% = dce_push_array(Socket,"day_yn_in",day_yn_in())
		rv% = dce_push_string(Socket,"user_in",user_in)
		rv% = dce_push_string(Socket,"place_in",place_in)
		Call dce_push_long(Socket,"count_in",count_in)
		dce_table = dce_submit("sl_p_01","sl_p_01_spcnum_sl",Socket)
	End If
	sl_p_01_spcnum_sl = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spcno",spcno())
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_01_spcnum_poct& (ptno_in$(),orddte_in$(),hspcl_in$(),frctcd_in$(),frctcd_extra_in$(),spccd_in$(),wkgrp_in$(),ordsite_in$(),dpcd_in$(),rsvdte_in$(),iogb_in$(),remk_in$(),job_in$(),day_yn_in$(),user_in$,place_in$,count_in&,spcno$(),msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"ptno_in",ptno_in())
		rv% = dce_push_array(Socket,"orddte_in",orddte_in())
		rv% = dce_push_array(Socket,"hspcl_in",hspcl_in())
		rv% = dce_push_array(Socket,"frctcd_in",frctcd_in())
		rv% = dce_push_array(Socket,"frctcd_extra_in",frctcd_extra_in())
		rv% = dce_push_array(Socket,"spccd_in",spccd_in())
		rv% = dce_push_array(Socket,"wkgrp_in",wkgrp_in())
		rv% = dce_push_array(Socket,"ordsite_in",ordsite_in())
		rv% = dce_push_array(Socket,"dpcd_in",dpcd_in())
		rv% = dce_push_array(Socket,"rsvdte_in",rsvdte_in())
		rv% = dce_push_array(Socket,"iogb_in",iogb_in())
		rv% = dce_push_array(Socket,"remk_in",remk_in())
		rv% = dce_push_array(Socket,"job_in",job_in())
		rv% = dce_push_array(Socket,"day_yn_in",day_yn_in())
		rv% = dce_push_string(Socket,"user_in",user_in)
		rv% = dce_push_string(Socket,"place_in",place_in)
		Call dce_push_long(Socket,"count_in",count_in)
		dce_table = dce_submit("sl_p_01","sl_p_01_spcnum_poct",Socket)
	End If
	sl_p_01_spcnum_poct = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spcno",spcno())
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_01_label& (spcno_in$,bar1$,bar2$,acptno&,msg$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"spcno_in",spcno_in)
		dce_table = dce_submit("sl_p_01","sl_p_01_label",Socket)
	End If
	sl_p_01_label = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_string(dce_table,"bar1",bar1)
	rv% = dce_pop_string(dce_table,"bar2",bar2)
	acptno = dce_pop_long(dce_table,"acptno")
	rv% = dce_pop_string(dce_table,"msg",msg)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_01_cancel& (spcno_in$(),ptno_in$,user_in$,orddte_in$(),msg$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"spcno_in",spcno_in())
		rv% = dce_push_string(Socket,"ptno_in",ptno_in)
		rv% = dce_push_string(Socket,"user_in",user_in)
		rv% = dce_push_array(Socket,"orddte_in",orddte_in())
		dce_table = dce_submit("sl_p_01","sl_p_01_cancel",Socket)
	End If
	sl_p_01_cancel = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_string(dce_table,"msg",msg)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_01_upd_day_yn& (spcno_in$(),day_yn_in$(),day_remk_in$(),msg$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"spcno_in",spcno_in())
		rv% = dce_push_array(Socket,"day_yn_in",day_yn_in())
		rv% = dce_push_array(Socket,"day_remk_in",day_remk_in())
		dce_table = dce_submit("sl_p_01","sl_p_01_upd_day_yn",Socket)
	End If
	sl_p_01_upd_day_yn = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_string(dce_table,"msg",msg)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_01_11_main& (frptno_in$,toptno_in$,frwdno_in$,towdno_in$,frctcd_in$,orddte_in$,hspcl_in$,user_in$,place_in$,spcno$(),msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"frptno_in",frptno_in)
		rv% = dce_push_string(Socket,"toptno_in",toptno_in)
		rv% = dce_push_string(Socket,"frwdno_in",frwdno_in)
		rv% = dce_push_string(Socket,"towdno_in",towdno_in)
		rv% = dce_push_string(Socket,"frctcd_in",frctcd_in)
		rv% = dce_push_string(Socket,"orddte_in",orddte_in)
		rv% = dce_push_string(Socket,"hspcl_in",hspcl_in)
		rv% = dce_push_string(Socket,"user_in",user_in)
		rv% = dce_push_string(Socket,"place_in",place_in)
		dce_table = dce_submit("sl_p_01","sl_p_01_11_main",Socket)
	End If
	sl_p_01_11_main = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spcno",spcno())
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_01_11_print& (frptno_in$,toptno_in$,frwdno_in$,towdno_in$,frspcno_in$,tospcno_in$,frctcd_in$,hspcl_in$,orddte_in$,spcno$(),msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"frptno_in",frptno_in)
		rv% = dce_push_string(Socket,"toptno_in",toptno_in)
		rv% = dce_push_string(Socket,"frwdno_in",frwdno_in)
		rv% = dce_push_string(Socket,"towdno_in",towdno_in)
		rv% = dce_push_string(Socket,"frspcno_in",frspcno_in)
		rv% = dce_push_string(Socket,"tospcno_in",tospcno_in)
		rv% = dce_push_string(Socket,"frctcd_in",frctcd_in)
		rv% = dce_push_string(Socket,"hspcl_in",hspcl_in)
		rv% = dce_push_string(Socket,"orddte_in",orddte_in)
		dce_table = dce_submit("sl_p_01","sl_p_01_11_print",Socket)
	End If
	sl_p_01_11_print = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spcno",spcno())
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


Function sl_p_01_12_main& (frwdno_in$,towdno_in$,orddte_in$,user_in$,place_in$,spcno$(),msg_out$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_01")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"frwdno_in",frwdno_in)
		rv% = dce_push_string(Socket,"towdno_in",towdno_in)
		rv% = dce_push_string(Socket,"orddte_in",orddte_in)
		rv% = dce_push_string(Socket,"user_in",user_in)
		rv% = dce_push_string(Socket,"place_in",place_in)
		dce_table = dce_submit("sl_p_01","sl_p_01_12_main",Socket)
	End If
	sl_p_01_12_main = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"spcno",spcno())
	rv% = dce_pop_string(dce_table,"msg_out",msg_out)
	Call dce_table_destroy(dce_table)
End function


