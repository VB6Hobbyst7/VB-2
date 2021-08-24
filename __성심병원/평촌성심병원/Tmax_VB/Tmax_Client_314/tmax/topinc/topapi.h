
/* ---------------------- topinc/topapi.h --------------------- */
/*								*/
/*              Copyright (c) 2000 Tmax Soft Co., Ltd		*/
/*                   All Rights Reserved  			*/
/*								*/
/* ------------------------------------------------------------ */

#ifndef _TOPEND_API_H
#define _TOPEND_API_H

#include <limits.h>

#define TP_USERID_LEN    12    /* length of dialogue user id         */
#define TP_PASSWORD_LEN  12    /* length of dialogue user Pass Word  */
#define TP_ENDPOINT_LEN  20    /* length of endpoint name            */
#define TP_PROD_NAME_LEN 32    /* length of product name             */
			       /* definition moved to tp_types.h     */
#define TP_FUNC_NAME_LEN  8    /* length of function name within prod*/
#define TP_FMT_NAME_LEN   8    /* length of format name within prod  */
#define TP_MSR_TARG_NAME_LEN 8 /* length of msr target name          */
#define TP_MCR_QUAL_LEN   8    /* no longer used: replaced by TP_MSR_TARG_LEN,
				  above. may be removed in future release. */
#define TP_NODE_ID_LEN    8    /* length of a internal TOP END node id */
#define TP_NODENAME_LEN   TP_NODE_ID_LEN    /* obsolete as of 2.04 */
#define TP_NODE_NAME_LEN  256  /* max length of an external node name */
#define TP_LANG_LEN       2    /* length of a format language code   */
#define TP_FMTTYPE_LEN    4    /* length of a format type code       */
#define TP_MODNAME_LEN   32    /* length of module message identifier*/
#define TP_FMTFILENAME_LEN 256 /* length of format file name         */
#define TP_SYSNAME_LEN    8    /* length of a TOP END system name    */
#define TP_MAX_BUF_LEN    (30 * 1024)   /* Max length of user buffer */


#define TP_ANY_DIALOGUE   -1    
#define TP_SUBORDINATE_DIALOGUE -2 
#define TP_UNIQUE_DIALOGUE -3      

#define TP_SET           0x00000001L 
#define TP_CLEAR         0x00000002L 

#define TP_MENU_PROD "TOPEND SYSTEM                   "
#define TP_MENU_FUNC "MENU    "


#define TP_BLOCK (long)LONG_MAX  /* suspend, wait for message    */
#define TP_NOBLOCK     0L         /* do not suspend               */
#define TP_INIT_TIME   20       /* suspend on startup for a 
                                    max of 20 seconds */


#define TP_NOFLAGS        0x00000000L /* no flags set                    */
#define TP_DISSOLVE       0x00000001L /* Dissolve dialogue on output     */
#define TP_APPL_CONTEXT   0x00000002L /* Application context exists      */
#define TP_NON_TRANSACT   0x00000004L /* Ignore Transaction Mode         */
#define TP_NO_RESPONSE    0x00000008L /* No Response only, no reply      */
#define TP_SIGNON_EVENT   0x00000010L /* Outcome from sign on request    */
#define TP_SIGNOFF_EVENT  0x00000020L /* Outcome from sign off request   */
#define TP_SIGNOFF_IMMED  0x00000040L /* Abort Dialogue Immediately      */
#define TP_ROLLBACK_ONLY  0x00000080L /* Client should rollback tx       */
#define TP_MSGTRUNC       0x00001000L /* Returning msg has been truncated*/
#define TP_SHUTTING_DOWN  0x00002000L /* Warn appl of shutdown in prog   */
#define TP_IGNORE_ADMIN   0x00004000L /* Do not return TP_ADMIN          */
#define TP_TRUNCATE       0x00008000L /* Truncate messages if necessary  */
#define TP_RESET_DIALOGUE 0x00010000L /* Abort subordinate dialogues and
                                         reset this dialogue             */
#define TP_CLIENT_MSG     0x00020000L /* Return from server receive if
                                         client message received         */
#define TP_TRANSERR       0x00040000L /* non-fatal translation error
                                         occurred.                       */
#define TP_ATTACHMENT     0x00080000L /* Attachment has been sent to     */
                                      /* this component and must be      */
                                      /* dispositioned.                  */


#define TP_DIF_DIAL_INFO     0x00000001L  /* tp_dialogue_info_t struct*/
#define TP_DIF_DIAL_USER     0x00000002L  /* tp_dialogue_user_t struct*/
#define TP_DIF_USER_INFO     0x00000004L  /* tp_user_info_t struct    */
#define TP_DIF_SERVICE_NAME  0x00000008L  /* tp_service_name_t struct */
#define TP_DIF_INPUT_FORMAT  0x00000010L  /* tp_input_format_t struct */
#define TP_DIF_OUTPUT_FORMAT 0x00000020L  /* tp_output_format_t struct*/
#define TP_DIF_LOCATION      0x00000040L  /* tp_location_t struct     */
#define TP_DIF_ATTACH_INFO   0x00000080L  /* tp_attach_info_t struct  */
#define TP_DIF_ALL           0x0FFFFFFFL  /* allocate all of the above*/


#define TP_ADMIN_EVENT       0x00000001L  /* Outstanding admin  msg */
#define TP_CSI_CLIENT_EVENT  0x00000002L  /* Outstanding client msg */
#define TP_CSI_SERVER_EVENT  0x00000004L  /* Outstanding server msg */


#define TP_ATTACH_TYPE_UNDEFINED  0   /* for clearing the field. */
#define TP_ATTACH_TYPE_FILE       1   /* standard file. */
#define TP_ATTACH_TYPE_PIPE       2   /* pipe.   */

#define TP_ATTACH_MAX_SIZE        2147483647
#define TP_ATTACH_PATH_LEN        256

#define TP_ATTACH_NO_FLAGS        0x00000000L  
#define TP_ATTACH_LOCAL_OPTIMIZE  0x00000001L  
#define TP_ATTACH_RTQ_OPTIMIZE    0x00000002L  
#define TP_ATTACH_TRANSIENT       0x00000004L  
#define TP_ATTACH_VALIDATE_SUM    0x00000010L  
#define TP_ATTACH_VALIDATE_FINAL  0x00000020L  

#define TP_ATTACH_TRANSFER        0x00000001L  
#define TP_ATTACH_CANCEL          0x00000002L  
#define TP_ATTACH_KEEP_ON_FAILURE 0x00000008L  



#define TP_OK                 0  
#define TP_DIFERR            -1  
#define TP_INTERR            TP_DIFERR
#define TP_RETURNED          -2  
#define TP_TIMEOUT           -3  
#define TP_SERVICE           -4  
#define TP_RESET             -5  
#define TP_DISSOLVED         -6  
#define TP_PROTOERR          -7  
#define TP_DIALOGUE          -8  
#define TP_INIT              -9  
#define TP_SYSTEM           -10  
#define TP_FMTFILE          -11  
#define TP_FIFERR           -12  
#define TP_FMT_NOTFOUND     -13  
#define TP_FMT_POINTER      -14  
#define TP_BUFFSIZE         -15  
#define TP_ADMIN            -16  
#define TP_NOMESSAGE        -17  
#define TP_DUPLICATE        -18  
#define TP_USER             -19  
#define TP_DISCONNECT       -20  
#define TP_MEMERR           -21  
#define TP_PARAMERR         -22  
#define TP_NAMEERR          -23  
#define TP_SIGNON_INHIBITED -24  
#define TP_SHUTDOWN         -25  
#define TP_FLAGERR          -26  
#define TP_NOT_SERVER       -27  
#define TP_DISSOLVING       -28  
#define TP_CLEAN_SHUTDOWN   -29  
#define TP_INIT_CALLED      -30  
#define TP_STOP             -31  
#define TP_USER_DATA_EXIT_ERROR -32 
#define TP_MESSAGE           -33    
#define TP_SMAPI_MESSAGE     -34    
#define TP_ATTACH_FLAGERR    -35    
#define TP_ATTACH_RECEIVED   -36    
#define TP_ATTACH_FAILED     -37    
#define TP_NOT_AVAIL         -99    
                                    
#define TP_EXT_MULT_RESP      101 
#define TP_EXT_FUNC_SWITCH    102 
#define TP_EXT_SIGNOFF        103 
#define TP_EXT_MULT_RECV      104 
#define TP_EXT_NO_LANG        105 
#define TP_EXT_SERVER_FAILED  106 
#define TP_EXT_APPL_REQUESTED 107 
                                  
                                  
#define TP_EXT_NO_FUNC_SWITCH 108 
#define TP_EXT_ROOT_TX        109 
#define TP_EXT_SUB_DLGS       110 
#define TP_EXT_CLIENT_FAILED  111 
#define TP_EXT_NO_RESP_ERR    112 
#define TP_EXT_SERVER_APPL    113 
#define TP_EXT_SUB_DLG_TX     114 
                                  
#define TP_EXT_SERVER_DLG_TX  115 
#define TP_EXT_SERVER_TX      116 
#define TP_EXT_SERVER_NOT_TX_OPENED 117  
#define TP_EXT_ROLLBACK       118 
#define TP_EXT_CLIENT_DIED    119 
#define TP_EXT_MULT_SEND      120 
#define TP_EXT_SERVER_TM_ROLLBACK 121 
#define TP_EXT_CLIENT_MSG     122 
#define TP_EXT_NO_SUCH_SERV   123 
#define TP_EXT_NOT_AUTHORIZED 124 
#define TP_EXT_ND_SHUTDOWN    125 
#define TP_EXT_MSR_FAILURE    126 
#define TP_EXT_MSR_KEY_LENGTH 127 
#define TP_EXT_MSR_NULL_TERM  128 
#define TP_EXT_MSR_INTERR     129 
#define TP_EXT_TRANSLATION    130 
#define TP_EXT_ROUTE_TIMEOUT  131 
#define TP_EXT_FEATURE_DISABLED  132 
#define TP_EXT_ATTACH_ACTIVE  133 
#define TP_EXT_ATTACH_TRANS   134 
#define TP_EXT_ATTACH_FWD     135 
#define TP_EXT_ATTACH_FWD_FAILURE   136 
#define TP_EXT_ATTACH_DISSOLVE_ERR  137 
                                        
                                        
typedef struct tp_sys_dialogue_id {
   char tp_nodename[TP_NODE_ID_LEN];  /* Dialogue source node identifier */
   unsigned short tp_sys_dialogue;    /* System dialogue specific id */
   unsigned short tp_ident_check;     /* For CSI internal use only   */
} tp_sys_dialogue_id_t;


typedef struct tp_dialogue_user  {
   char tp_userid[TP_USERID_LEN];     /* User identification       */
   char tp_password[TP_PASSWORD_LEN]; /* User's Kerberos Pass Word */
   char tp_endpoint[TP_ENDPOINT_LEN]; /* User end point name       */
} tp_dialogue_user_t;


typedef struct tp_dialogue_info  {
   long tp_user_dialogue_id;          /* User assigned dialogue id  */
   long tp_user_message_id;           /* User assigned message id   */
   tp_sys_dialogue_id_t tp_system_dialogue_id;
                                      /* System assigned dialogue id*/
   long tp_flags;                     /* Processing flags           */
   long tp_extended_status;           /* Extended status value      */
} tp_dialogue_info_t;


typedef struct tp_user_info  {
   char tp_userid[TP_USERID_LEN];     /* User identification          */
   char tp_endpoint[TP_ENDPOINT_LEN]; /* User end point name          */
   long tp_time_stamp;                /* Message origination timestamp*/
} tp_user_info_t;


typedef struct tp_service_name  {
   char tp_product_name[TP_PROD_NAME_LEN];  /* Required - name of product 
                                               in which specific function 
                                               is requested            */
   char tp_function_name[TP_FUNC_NAME_LEN]; /* Required - requested 
                                               function name           */
   long tp_function_qualifier;              /* optional - message code */
   char tp_message_routing_qualifier[TP_MSR_TARG_NAME_LEN]; 
} tp_service_name_t;


typedef struct tp_system_info  {
   char tp_system[TP_SYSNAME_LEN];    /* TOP END system name            */
   char tp_product[TP_PROD_NAME_LEN]; /* Component's Product Name       */
   char tp_node[TP_NODE_ID_LEN];      /* Node identifier                */
   long tp_process_id;                /* Component's Process ID         */
   long tp_dialogue_count;            /* Number active in this component*/
} tp_system_info_t;


typedef struct tp_location {
    char tp_node_name[TP_NODE_ID_LEN]; /* TOP END node identifier */
} tp_location_t;

typedef struct tp_attach_id_struct {
   long tp_time_val;
   long tp_attach_id;
} tp_attach_id_t;

typedef struct tp_attach_info {
   tp_attach_id_t  tp_attach_id;
   long   tp_attach_flags; 
   long   tp_attach_type;  
   long   tp_attach_size;  
   union {
      char   tp_attach_path[TP_ATTACH_PATH_LEN]; 
   } tp_attach_location;
   long   tp_reserved[2];            
} tp_attach_info_t;

typedef struct tp_funct_struct {
   char tp_function_name[TP_FUNC_NAME_LEN]; 
} tp_funct_struct_t;


typedef struct tp_application_info {
   char tp_product_name[TP_PROD_NAME_LEN];  
   char tp_format_file[TP_FMTFILENAME_LEN]; 
} tp_application_info_t;


typedef struct tp_input_format  {
   char tp_format_name[TP_FMT_NAME_LEN];  
   char tp_format_language[TP_LANG_LEN];  
   char tp_format_type[TP_FMTTYPE_LEN];   
} tp_input_format_t;


typedef struct tp_output_format  {
   char tp_format_name[TP_FMT_NAME_LEN];  
   long tp_format_qualifier;              
   long tp_format_revision;               
} tp_output_format_t;


typedef struct tp_retrieve_format  {
   char tp_product_name[TP_PROD_NAME_LEN]; 
   char tp_format_name[TP_FMT_NAME_LEN];   
   char tp_format_language[TP_LANG_LEN];   
   char tp_format_type[TP_FMTTYPE_LEN];    
   char pad[2];                            
   long tp_format_revision;                
} tp_format_name_t;

#define tp_format_header_t	tp_fmt_key_t
#define tp_fm_key_t		tp_fmt_key_t


typedef struct tp_dif_structs {
   tp_dialogue_info_t   *info;
   tp_dialogue_user_t   *client;
   tp_user_info_t       *source_info;
   tp_service_name_t    *service;
   tp_input_format_t    *input_format;
   tp_output_format_t   *output_format; 
   tp_location_t        *location;
   tp_attach_info_t     *attach_info;
} tp_dif_structs_t;


#ifdef __cplusplus
extern "C" {
#endif

#if defined(__STDC__) || defined(__cplusplus) || defined(_WIN32)
extern tp_dif_structs_t *tp_csi_alloc(unsigned long fields);
extern int tp_csi_free(tp_dif_structs_t *mem);
extern int tp_client_signoff(tp_dialogue_info_t *info);
extern int tp_initialize(tp_application_info_t *application_info,
                         tp_funct_struct_t     function_array[],
                         long                  number_functions);

extern int tp_terminate(void);
extern int tp_system_info(tp_system_info_t *system_info);
extern int tp_system_admin(long *flags,
                           long *buffer_length,
                           char *buffer_text);
extern int tp_system_message(char *sys_message_number,
                             char *sys_message_data,
                             char *module_name,
                             int   module_line_number);
extern int tp_receive_end(long mode);
extern void tp_csi_refresh(tp_dif_structs_t *dif_ptr,
                           unsigned long fields);
extern int tp_csi_size(unsigned long field);
extern char *tp_csi_status_name(int status);
extern int tp_daemonize(long always);
extern int tp_log_signals();

#if TP_API_VERSION == 2
extern int tp_client_signon_v2(tp_dialogue_info_t *info,
                            tp_dialogue_user_t *client,
                            long               inactivity_time,
                            tp_service_name_t  *service,
                            tp_input_format_t  *input_format,
                            long               message_length,
                            char               *message_text,
                            tp_attach_info_t   *attach_info);

#define tp_client_signon tp_client_signon_v2

extern int tp_client_receive_v2(tp_dialogue_info_t *info,
                             long               wait_time,
                             tp_output_format_t *output_format,
                             tp_service_name_t  *service,
                             tp_location_t      *location,
                             long               *message_length,
                             char               *message_text,
                             tp_attach_info_t   *attach_info);

#define tp_client_receive tp_client_receive_v2

extern int tp_client_send_v2(tp_dialogue_info_t *info,
                          tp_service_name_t  *service,
                          tp_input_format_t  *input_format,
                          long               message_length,
                          char               *message_text,
                          tp_attach_info_t   *attach_info);

#define tp_client_send tp_client_send_v2

extern int tp_server_receive_v2(tp_dialogue_info_t *info,
                             long               wait_time,
                             tp_service_name_t  *service,
                             tp_input_format_t  *input_format,
                             tp_location_t      *location,
                             tp_user_info_t     *source_info,
                             long               *message_length,
                             char               *message_text,
                             tp_attach_info_t   *attach_info);

#define tp_server_receive tp_server_receive_v2

extern int tp_server_send_v2(tp_dialogue_info_t *info,
                          tp_output_format_t *output_format,
                          long               message_length,
                          char               *message_text,
                          tp_attach_info_t   *attach_info);

#define tp_server_send tp_server_send_v2

extern int tp_process_attachment(tp_dialogue_info_t *info,
                                 long               action_flags,
                                 tp_attach_info_t   *attach_info);

#else    /* else if TP_API_VERSION != 2 */

extern int tp_client_signon(tp_dialogue_info_t *info,
                            tp_dialogue_user_t *client,
                            long               inactivity_time,
                            tp_service_name_t  *service,
                            tp_input_format_t  *input_format,
                            long               message_length,
                            char               *message_text);

extern int tp_client_receive(tp_dialogue_info_t *info,
                             long               wait_time,
                             tp_output_format_t *output_format,
                             tp_service_name_t  *service,
                             tp_location_t      *location,
                             long               *message_length,
                             char               *message_text);

extern int tp_client_send(tp_dialogue_info_t *info,
                          tp_service_name_t  *service,
                          tp_input_format_t  *input_format,
                          long               message_length,
                          char               *message_text);

extern int tp_server_receive(tp_dialogue_info_t *info,
                             long               wait_time,
                             tp_service_name_t  *service,
                             tp_input_format_t  *input_format,
                             tp_location_t      *location,
                             tp_user_info_t     *source_info,
                             long               *message_length,
                             char               *message_text);

extern int tp_server_send(tp_dialogue_info_t *info,
                          tp_output_format_t *output_format,
                          long               message_length,
                          char               *message_text);
#endif   /* end if TP_API_VERSION == 2 */

#else

extern tp_dif_structs_t *tp_csi_alloc();
extern int tp_csi_free();
extern int tp_client_receive();
extern int tp_client_send();
extern int tp_client_signon();
extern int tp_client_signoff();
extern int tp_server_receive();
extern int tp_server_send();
extern int tp_initialize();
extern int tp_terminate();
extern int tp_system_info();
extern int tp_system_admin();
extern int tp_system_message();
extern int tp_receive_end();
extern void tp_csi_refresh();
extern int tp_csi_size();
extern char *tp_csi_status_name();
extern int tp_daemonize();
extern int tp_log_signals();
#endif

#ifdef __cplusplus
}
#endif

#endif   /* _TOPEND_API_H */
