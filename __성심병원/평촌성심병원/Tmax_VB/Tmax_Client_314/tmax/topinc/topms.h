
/* ------------------------ topinc/topms.h -------------------- */
/*								*/
/*              Copyright (c) 2000 Tmax Soft Co., Ltd		*/
/*                   All Rights Reserved  			*/
/*								*/
/* ------------------------------------------------------------ */

#ifndef _TOPEND_MS_H
#define _TOPEND_MS_H


#define		COBOL_INT	11	
#define		COBOL_LONG	11	

#define		COBOL_SHORT	 6	

#define		TP_EXT_SYSTEM_ERROR 		101		
#define		TP_EXT_FUNCTION_INITIALIZATION 	201		
#define		TP_EXT_REQ_ERROR 		202		
#define		TP_APPLICATION_ERROR 		204		

#define		TP_EXT_FQ_UNDEFINED 		203		

typedef struct tp_inf_appl_status{

	char 	appl_status[COBOL_LONG];		
	char	appl_extended_status[COBOL_LONG];		

}tp_inf_appl_status_t;

typedef struct tp_inf_service_name{

	char	service_product_name[TP_PROD_NAME_LEN];		
	char	service_function_name[TP_FUNC_NAME_LEN];		
	char	service_function_qualifier[COBOL_LONG];

}tp_inf_service_name_t;

typedef struct tp_inf_format{

	char	format_name[TP_FMT_NAME_LEN];
	char	format_language[TP_LANG_LEN];
	char	format_type[TP_FMTTYPE_LEN];
	char	format_qualifier[COBOL_LONG];

}tp_inf_format_t;


typedef struct tp_inf_source{

	char	source_date_year[4];
	char	source_date_month[2];
	char	source_date_day[2];
	char	source_time_hour[2];
	char	source_time_minutes[2];
	char	source_time_seconds[2];
	char	source_time_hundreds[2];
	char	source_endpoint[TP_ENDPOINT_LEN];
	char	source_user_id[TP_USERID_LEN];
	char	source_node_name[TP_NODE_ID_LEN];

}tp_inf_source_t;

typedef struct tp_inf_req{

	char	req_disconnect_notification;
	char	req_rollback_only;
	char	req_reset_dialogue;
	char	req_dissolve_dialogue;
	char	req_goto_exit;		/* goto_exit feature */
	char	req_future_use_area[9];

}tp_inf_req_t;

typedef struct tp_inf_ind{

	char	ind_dissolve_dialogue;
	char	ind_no_response;
	char	ind_input_truncated;
	char	ind_future_use_area[10];
	
}tp_inf_ind_t;

/*==============================================================*
 *	TOP END Managed Server C data structures		*
 *==============================================================*/

typedef struct tp_ms_inf{
	char			inf_layout_revision[2];
	tp_inf_appl_status_t	inf_status;
	tp_inf_service_name_t	inf_service;
	tp_inf_format_t		inf_format;
	tp_inf_source_t		inf_source;
	tp_inf_req_t		inf_req;
	tp_inf_ind_t		inf_ind;

}tp_ms_inf_t;

/*====================================================================*
 *	Managed Server Argument Definition			      *
 *====================================================================*/

typedef struct tp_ms_area{ 
	char	area_length[COBOL_LONG];
	char	area_fill;
	char	data[1];
}tp_ms_area_t;


#endif /* _TOPEND_MS_H */
