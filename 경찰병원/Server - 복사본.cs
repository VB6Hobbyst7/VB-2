using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.Odbc;
using Oracle.DataAccess.Client;
using Microsoft.VisualBasic;
using CnsUtils;
using DataModule;

namespace HostDLL
{
    public class ServerDLL
    {
        //경찰병원
        public static string SERVER_ID = "";
        public static string Text_DNS = "";
        public string err_log = "";

        string tns_ora=string.Format(@"
         DESCRIPTION =
          (ADDRESS_LIST =
             (ADDRESS = (PROTOCOL = TCP)(HOST =10.60.210.14)(PORT = 1521))
             (ADDRESS = (PROTOCOL = TCP)(HOST =10.60.210.13)(PORT = 1521))
          )
         (CONNECT_DATA =
              (SERVICE_NAME = NPHHISP)
              (SERVER = DEDICATED)
          )
         ");
     
        sex_age m_sex_age= new sex_age();
        string conn_string = "";
        public ServerDLL()
        {
            conn_string = string.Format("Data Source=({0});User Id=lis_inf;Password=lisprod", tns_ora);
        }


        public MyData.orderDataTable Get_NoAcpt_Spc(string spc_no,string option)
        {
            MyData.orderDataTable order = new MyData.orderDataTable();
            return order;
        }
       
        public string GetPatInfo(string pat_no)
        {
            string pat_info = "";

            return pat_info;
        }
        public string GetUserName(string user_id)
        {
            string user_nm = "";

            return user_nm;
        }
        public string Check_User(string user_id,string password)
        {
            string user_nm = user_id;
            try
            {
                string query = string.Format(@"SELECT USER_NM,PW,ECL_DIGEST('{1}') 
                     FROM AZCMMUSER WHERE USID='{0}'", user_id, password);
                DataTable dt_ret = Select(query);
                if (dt_ret != null || dt_ret.Rows.Count > 0)
                {
                    user_nm = dt_ret.Rows[0]["USER_NM"].ToString();
                }
            }
            catch(Exception ex)
            {

            }
            return user_nm;
        }

        public void UpdateEvent(string spc_no, string pat_no, string user_id, string equp_cd)
        {

        }
        public string UploadFileToFtpServer(string file_path, string fiel_nm, byte[] file_stream)
        {
            string url = "";

            return url;
        }

        public MyData.orderDataTable Download(
            string spc_no,
            string work_cd,
            List<string> item_list,
            string option
            )
        {

            MyData.orderDataTable order = new MyData.orderDataTable();
            if (spc_no == "") return order;

            string exam_code_list = "";
            if (item_list != null)
            {
                foreach (string item in item_list)
                {
                    exam_code_list += "'" + item.Trim() + "',";
                }

            }
            string query = string.Format(@"
               SELECT DISTINCT
                       J1.SPCM_NO
                     , J2.BRCD_LABL_NO AS SPC_NO
                     , J1.PID
                     , J1.PT_NM
                     , J1.SEX
                     , J1.AGE
                     , R1.RCPN_DT 
                     , J1.DOBR AS BIRTH
                     , TO_CHAR(R1.RCPN_DT,'yyyyMMdd') AS ACPT_DT
                     , R1.RSLT_NO AS ACPT_NO
                     , J1.SLIP_CD 
                     , J1.SPCM_CD 
                     , J1.PID
                     , J1.PT_NM
                     , J1.SEX
                     , J1.AGE
                     , R1.RCPN_DT 
                     , J1.DOBR AS BIRTH
                     , TO_CHAR(R1.RCPN_DT,'yyyyMMdd') AS ACPT_DT
                     , R1.RSLT_NO AS ACPT_NO
                     , J1.SLIP_CD 
                     , J1.SPCM_CD 
                     , R1.EXMN_CD AS ITEMCD
                     , D1.ENGL_CD AS DEPTCD                
                   FROM 
                      SPSLMJBDI J2
                     ,SPSLHRRST R1
                     ,SPSLMJBBI J1
                     ,AZCMMDEPT D1 
                     ,SPSLMFBIF F1
                WHERE 
                  J2.BRCD_LABL_NO='{0}'
          AND J2.SPCM_NO = R1.SPCM_NO
          AND J1.SPCM_NO = J2.SPCM_NO 
          AND J1.MED_DP = D1.DEPT_CD 
          AND R1.EXMN_CD = F1.EXMN_CD  (+)
           AND R1.RCPN_DT >= F1.USE_STR_DY  (+) 
                   AND R1.RCPN_DT < F1.USE_END_DY  (+)
                   AND F1.MNDT_YN <> 'Y' ", spc_no);
            if (exam_code_list != "")
                query += " AND R1.EXMN_CD IN (" + exam_code_list.Substring(0, exam_code_list.Length - 1) + ")";

            string query_qc = string.Format(@"
               SELECT 
                       '' SPCM_NO
                     , SPCM_NO AS SPC_NO
                     , SBSN_CD AS PID
                     , LOT_NO AS PT_NM
                     , '' AS SEX
                     , '' AS AGE
                     , TO_CHAR(SYSDATE,'yyyyMMdd') AS ACPT_DT
                     , RSLT_SQNO AS ACPT_NO
                     , '' AS SLIP_CD 
                     , '' AS SPCM_CD 
                     , EXMN_CD AS ITEMCD
                     , '' AS DEPTCD 
               FROM  SPSLHQRST
                    WHERE SPCM_NO = '{0}' 
               AND DEL_YN = 'N'
               AND (RSLT_VALU is null OR RSLT_VALU = '')", spc_no);
           
            DataTable dt_ret = new DataTable();
            try
            {

                if (spc_no.Substring(2, 1) == "9")
                {
                    SysFunc.SaveSystemLog(query_qc);
                    dt_ret = Select(query_qc);
                }
                else
                {
                    SysFunc.SaveSystemLog(query);
                    dt_ret = Select(query);
                }
                foreach (DataRow row in dt_ret.Rows)
                {
                    MyData.orderRow order_row = order.NeworderRow();
                    order_row.spc_no = spc_no;
                    order_row.pat_no = row["PID"].ToString();
                    order_row.pat_nm = row["PT_NM"].ToString();
                    order_row.age = row["AGE"].ToString();
                    order_row.sex = row["SEX"].ToString();
                    order_row.dept = row["DEPTCD"].ToString();
                    order_row.exam_cd = row["ITEMCD"].ToString();
                    order_row.acpt_dt = row["ACPT_DT"].ToString();
                    order_row.acpt_no = row["ACPT_NO"].ToString();
                    order_row.lab_info = row["SPCM_NO"].ToString();
                    order_row.spc_cd = row["SPCM_CD"].ToString();
                    order_row.qc_yn = "";
                    order.Rows.Add(order_row);
                }
            }
            catch (Exception ex)
            {
                SysFunc.SaveSystemLog(string.Format("Select {0}", ex.Message) + "--->" + query);
            }

           
            return order;
        }

        private void UpdateStatus( 
            string spc_no,
            string stat,
            List<string> item_list)
        {
        }

        
         public MyData.orderDataTable Worklist(
             DateTime from_date,
             DateTime to_date,
             string work_cd,
             List<string> item_list,
             string option = "")
         {
             string exam_code_list = "";
             if (item_list != null)
             {
                 foreach (string item in item_list)
                 {
                     exam_code_list += "'" + item.Trim() + "',";
                 }

             }
             MyData.orderDataTable order = new MyData.orderDataTable();

             DataTable dt_ret = new DataTable();

             string query = string.Format(@"
                SELECT DISTINCT 
                       J2.BRCD_LABL_NO AS SPC_NO
                     , J1.PID AS UnitNo
                     , J1.SPCM_NO
                     , J1.PT_NM AS PatNm
                     , J1.SEX
                     , J1.AGE
                     , J1.DOBR AS BIRTH
                     , R1.RCPN_DT AS INPDTE
                     , R1.RSLT_NO AS WrkSeq
                     , J1.SPCM_CD AS SpcNm
                     , R1.PRSC_CD AS ITEMCD
                     , D1.ENGL_CD AS DEPTCD 
                  FROM SPSLMJBBI J1
                     , SPSLMJBDI J2
                     , SPSLHRRST R1
                     , SPSLMFBIF F1
                     , AZCMMDEPT D1 
                 WHERE J1.SPCM_NO = J2.SPCM_NO 
                   AND R1.RSLT_STAT <> '3'
                   AND J1.MED_DP = D1.DEPT_CD 
                   AND J2.SPCM_NO = R1.SPCM_NO
                   AND R1.EXMN_CD = F1.EXMN_CD  (+)
                   AND R1.RCPN_DT >= F1.USE_STR_DY  (+) 
                   AND R1.RCPN_DT < F1.USE_END_DY  (+)
                   AND R1.RCPN_DT BETWEEN to_date('{0}'||'000000', 'YYYYMMDDHH24MISS') AND to_date ('{1}'||'235959', 'YYYYMMDDHH24MISS')
                   AND F1.MNDT_YN <> 'Y'
                   AND J1.SLIP_CD in ({2})  "
                 , from_date.ToString("yyyyMMdd"), to_date.ToString("yyyyMMdd"), work_cd);
             if (exam_code_list != "")
                 query += "AND SUBSTR(R1.PRSC_CD, 1, 6) IN (" + exam_code_list.Substring(0, exam_code_list.Length - 1) + ")";
             query += " ORDER BY INPDTE, J1.PT_NM  ";

             dt_ret = Select(query);

             foreach (DataRow row in dt_ret.Rows)
             {
                 MyData.orderRow order_row = order.NeworderRow();
                 order_row.spc_no = row["SPC_NO"].ToString();
                 order_row.pat_no = row["PID"].ToString();
                 order_row.pat_nm = row["PT_NM"].ToString();
                 order_row.age = row["AGE"].ToString();
                 order_row.sex = row["SEX"].ToString();
                 order_row.dept = row["DEPTCD"].ToString();
                 order_row.exam_cd = row["ITEMCD"].ToString();
                 order_row.acpt_dt = row["ACPT_DT"].ToString();
                 order_row.acpt_no = row["ACPT_NO"].ToString();
                 order_row.lab_info = row["SPCM_NO"].ToString();
                 order_row.spc_cd = row["SPCM_CD"].ToString();
                 order_row.qc_yn = "";
                 order_row.qc_yn = "";
                 order.Rows.Add(order_row);
             }
             return order;
         }

        public int UploadResult(
            string spc_no,
            string rack_pos,
           DataRow[] data_list, 
            string equp_cd, 
            string user_id_1,
            string user_id_2=""
            )
        {
      
            string exam_cd = "", result = "", norm = "", lab_seq_no = "", panic = "", exam_tm = "", pat_no="";
            string rst_cd = "", inst_rst = "", inst_int = "", flag = "";
            int n_ret = 0;

            foreach (DataRow rr in data_list)
            {
                pat_no = rr["pat_no"].ToString();
                exam_cd = rr["exam_cd"].ToString();
                result = rr["result"].ToString();
                rst_cd = rr["rst_cd"].ToString();
                inst_rst = rr["inst_rst"].ToString();
                inst_int = rr["inst_int"].ToString();
                flag = rr["flag"].ToString();
                lab_seq_no = rr["lab_seq_no"].ToString();
                n_ret += UpdateResult(
                    spc_no, 
                    rack_pos,
                    pat_no, 
                    exam_cd,
                    rst_cd,
                    result,
                    inst_rst,
                    inst_int,
                    flag,
                    norm, 
                    panic, 
                    exam_tm, 
                    lab_seq_no,
                    user_id_1,
                    user_id_2, 
                    equp_cd);
            }
            return n_ret;
        }

        public int AutoReg(string spc_no,string user_id)
        {
            int n_ret = 0;
            return n_ret;
        }


        /// <summary>
        /// 일반검사 결과 갱신
        /// </summary>
        /// <param name="spc_no"></param>
        /// <param name="ord_code"></param>
        /// <param name="result"></param>
        /// <param name="equp_cd"></param>
        /// <param name="exam_id"></param>
        /// <returns></returns>
        /// UpdateResult(spc_no, lab_seq_no, pat_no, exam_cd, result, norm, panic, exam_tm, equp_cd, user_id);
        public int UpdateResult(
            string spc_no,
            string rack_pos,
            string pat_no,
            string exam_cd,
            string rst_cd,
            string result,
            string inst_rst,
            string inst_int,
            string flag,
            string norm,
            string panic,
            string exam_tm,
            string lab_seq_no,
            string user_id_1,
            string user_id_2,
            string equp_cd
            )
        {
            int n_ret = 0;
            if (result == "" || spc_no.Length < 3) return 0;
            if (user_id_1 == "")
            {
                user_id_1 = "20022008";
            }


            if (spc_no.Substring(2, 1) == "9")
            {
                string query_qc = string.Format(@"
                    UPDATE SPSLHQRST SET
                       RSLT_VALU = '{3}', 
                       RSLT_DT = SYSDATE, 
                       UPDT_DT = SYSDATE
                    WHERE 
                        SPCM_NO = '{0}'
                    AND SBSN_CD= '{1}'
                    AND EXMN_CD = '{2}'  "
              , spc_no, pat_no, exam_cd, result);
                n_ret = ExecSQL(query_qc);
                return n_ret;
            }

            string query = string.Format(@" 
                UPDATE SPSLHRRST SET 
                     REAL_RSLT = '{2}'
                   , VIEW_RSLT = '{2}'
                   , RSLT_INPS_ID = '{3}'
                   , AMEN_ID = '{3}'
                   , RSLT_INPT_DT = SYSDATE
                   , EXMN_EQPM = '{4}'
                   , EXMN_RCPN_NO='{5}'  
                   , UPDT_DT = SYSDATE  
                   , RSLT_NO = 1
                   , RSLT_STAT  = '1' 
                   , REEXAMYN = 'N'   
                 WHERE 
                       SPCM_NO = FN_LABCVTBCNO(NVL('{0}', '-'))  
                   AND EXMN_CD = '{1}' 
                   AND RSLT_STAT IN ('0','1','4') ",
                   spc_no, exam_cd.Trim(), result, user_id_1, equp_cd, rack_pos);
            n_ret = ExecSQL(query);

            string query_stat = string.Format(@"
                UPDATE SPSLMJBBI SET 
                   RSLT_STAT = 1, AMEN_ID  = '{1}', UPDT_DT  = SYSDATE 
                WHERE
                  SPCM_NO = FN_LABCVTBCNO(NVL('{0}', '-'))  
                  AND RSLT_STAT <> '3'  AND SPCM_STAT <= '3' ", spc_no, user_id_1);
            
            n_ret += ExecSQL(query_stat);
            return n_ret;
        }
        public DataTable Select(string query)
        {
            err_log = "";
            DataTable dt_ret = new DataTable();

            using (OracleConnection conn = new OracleConnection(conn_string))
            {
                try
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }
                    OracleDataAdapter adaptor = new OracleDataAdapter(query, conn);
                    adaptor.Fill(dt_ret);
                }
                catch (Exception ex)
                {
                    err_log = string.Format("Select {0}", ex.Message) + "--->" + query;
                    SysFunc.SaveSystemLog(err_log);
                }
            }
            return dt_ret;
        }

        public int ExecSQL(string query)
        {
            int n_ret = 0;
            err_log = "";
            using (OracleConnection conn = new OracleConnection(conn_string))
            {
                OracleCommand command = new OracleCommand();
                command.Connection = conn;
                command.CommandTimeout = 1;
                command.CommandType = CommandType.Text;

                try
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }
                    SysFunc.SaveSystemLog( query);
                    command.CommandText = query;
                    n_ret = command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    err_log = string.Format("ExecSQL {0}:{1}", ex.Message, query);
                    SysFunc.SaveLog("["+DateTime.Now.ToString("HH:mm:ss")+"]" + err_log + Environment.NewLine + query);
                }
            }
            return n_ret;
        }


        public int Upload_Bact_Rst(
             string spc_no,
             string org_no,
             string org_cd,
             string org_nm,
             string acpt_dt,
             string acpt_no,
             string exam_cd,
             DataTable tb_result,
             string lab_seq_no,
             string comment,
             string equp_cd,
            string user_id)
        {
            int n_ret = 0;
            string hosp_ord_cd = GetHospOrgCode(org_cd);
            n_ret = UploadBact(spc_no, org_no, hosp_ord_cd, org_nm, equp_cd);
            DeleteAnti(spc_no, org_no);

            foreach (DataRow row in tb_result.Rows)
            {
                UploadAnti(
                    spc_no,
                    org_no,
                    row["anti_cd"].ToString(),
                    row["anti_nm"].ToString(),
                    row["mic"].ToString(),
                    row["ris"].ToString(),
                    row["card_cd_a"].ToString());

            }
            UpdateHostStat(lab_seq_no, user_id, "M");


            return n_ret;
        }

        public int UploadBact(
                 string spc_no,
                 string org_no,
                 string org_cd,
                 string org_nm,
                  string equp_cd
            )
        {
            int n_ret = 0;
            string query = "";
            string query1 = "SELECT BREBARNUM FROM L_BREINF WHERE BREBARNUM='" + spc_no + "' AND BRESEQ=" + org_no + "";
            DataTable dt_check = Select(query1);
            if (dt_check.Rows.Count > 0)
            {
                query = "UPDATE L_BREINF SET BREBATCOD='" + org_cd + "', BREBATNAM='" + org_nm + "' ";
                query += " WHERE BREBARNUM='" + spc_no + "' AND BRESEQ=" + org_no + " ";
            }
            else
            {
                query = "INSERT INTO L_BREINF (BREBARNUM ,BRESEQ ,BREBATCOD ,BREBATNAM ,BRETEQUCOD,BREUPDUID ,BREUPDDTF ) VALUES ";
                query += "( '" + spc_no + "'," + org_no + ", '" + org_cd + "','" + org_nm + "','" + equp_cd + "','',TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS') )";
            }
            n_ret = ExecSQL(query);
            return n_ret;
        }

        public int DeleteAnti(
            string spc_no,
            string org_no
            )
        {
            int n_ret = 0;
            string query = "";
            query = "DELETE FROM L_AREINF WHERE  AREBARNUM='" + spc_no + "' AND AREBRESEQ='" + org_no + "' ";
            n_ret = ExecSQL(query);
            return n_ret;
        }

        public int UploadAnti(
            string spc_no,
            string org_no,
            string anti_cd,
            string anti_nm,
            string mic,
            string ris,
            string card_cd
            )
        {
            int n_ret = 0;
            string query = "";
            query = "INSERT INTO L_AREINF (AREBARNUM,AREBRESEQ,AREANTCOD,AREANTNAM,AREMICRST,ARERISRST,AREANTCRD) VALUES ";
            query += "( '" + spc_no + "'," + org_no + ", '" + anti_cd + "','" + anti_nm + "','" + mic + "','" + ris + "','" + card_cd + "' )";
            n_ret = ExecSQL(query);
            return n_ret;
        }

        public void UpdateHostStat(string lab_seq_no, string user_id, string stat = "M")
        {
            string query = string.Format(@"
             UPDATE msystechhis.S_AdrInf  
             SET
                  AdrStaTyp = 'M',
                  AdrActUid ='{1}', AdrActDtf= to_char(sysdate,'yyyyMMddHH24miss')
             WHERE 
                  ADRKEY = '{0}' ", lab_seq_no, user_id);
            ExecSQL(query);
        }

        public string GetHospOrgCode(string org_cd)
        {
            string hosp_org_cd = org_cd;
            string query = "SELECT BATCOD FROM L_BATMST WHERE BATREFCMT='" + org_cd + "'";
            DataTable dt_ret = Select(query);
            if (dt_ret.Rows.Count > 0)
            {
                hosp_org_cd = dt_ret.Rows[0]["BATCOD"].ToString();
            }
            return hosp_org_cd;
        }

        public bool CheckLogin(string user_id, ref string pass_word, ref string user_name)
        {
            bool is_pass = false;
            //using (SqlConnection conn = new SqlConnection(conn_string))
            //{
            //    SqlCommand command = new SqlCommand("UP_H7LIS_EMPL_R", conn);
            //    command.CommandType = CommandType.StoredProcedure;
            //    command.Parameters.Add("@ID", SqlDbType.VarChar).Value = user_id;

            //    SqlParameter par_name = new SqlParameter("@NAME", SqlDbType.VarChar, 1000);
            //    par_name.Direction = ParameterDirection.Output;
            //    command.Parameters.Add(par_name);
            //    SqlParameter par_pwd = new SqlParameter("@PW", SqlDbType.VarChar, 1000);
            //    par_pwd.Direction = ParameterDirection.Output;
            //    command.Parameters.Add(par_pwd);
            //    try
            //    {
            //        conn.Open();
            //        command.ExecuteNonQuery();
            //        temp_pwd = par_pwd.Value.ToString();
            //        Byte[] bytes = Encoding.UTF8.GetBytes(temp_pwd);
            //        pass_word = Encoding.UTF8.GetString(bytes);
            //        user_name = par_name.Value.ToString();
            //        if (user_name != "")
            //        {
            //            is_pass = true;
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        CnsUtils.SysFunc.SaveSystemLog(ex.Message);
            //    }
            //}

            return is_pass;
        }

        public string EncodingPW(string pass_word)
        {
            string temp_pwd = "";
            //using (SqlConnection conn = new SqlConnection(conn_string))
            //{
            //    SqlCommand command = new SqlCommand("UP_H7LIS_CONVERT_PW_R", conn);
            //    command.CommandType = CommandType.StoredProcedure;
            //    command.Parameters.Add("@PW", SqlDbType.VarChar).Value = pass_word;

            //    SqlParameter par_pwd = new SqlParameter("@CONVERT_PW", SqlDbType.VarChar, 1000);
            //    par_pwd.Direction = ParameterDirection.Output;
            //    command.Parameters.Add(par_pwd);
            //    try
            //    {
            //        conn.Open();
            //        command.ExecuteNonQuery();
            //        temp_pwd = par_pwd.Value.ToString();
            //        Byte[] bytes = Encoding.UTF8.GetBytes(temp_pwd);
            //        temp_pwd = Encoding.UTF8.GetString(bytes);
            //    }
            //    catch (Exception ex)
            //    {
            //        CnsUtils.SysFunc.SaveSystemLog(ex.Message);
            //    }
            //}

            return temp_pwd;
        }

        public DataTable GetEqupList()
        {
            string query = string.Format(@"
                SELECT	*
                  FROM	LIS_EQP_LIST_V");
            DataTable dt_ret = new DataTable();

            //using (SqlConnection conn = new SqlConnection(conn_string))
            //{
            //    SqlDataAdapter adpater = new SqlDataAdapter(query, conn);
            //    try
            //    {
            //        adpater.Fill(dt_ret);
            //    }
            //    catch (Exception ex)
            //    {
            //        SysFunc.SaveSystemLog(string.Format("{0}-> {1}", query, ex.Message));
            //    }
            //}
            return dt_ret;
        }

        public DataTable GetLabExam(string lab_cd = "")
        {
            string query = string.Format(@"
                SELECT	*
                  FROM	LIS_ORD_LIST_V");
            if (lab_cd != "")
            {
                query += " WHERE WORK_GB='" + lab_cd + "'";
            }
            DataTable dt_ret = new DataTable();

            //using (SqlConnection conn = new SqlConnection(conn_string))
            //{
            //    SqlDataAdapter adpater = new SqlDataAdapter(query, conn);
            //    try
            //    {
            //        adpater.Fill(dt_ret);
            //    }
            //    catch (Exception ex)
            //    {
            //        SysFunc.SaveSystemLog(string.Format("{0}-> {1}", query, ex.Message));
            //    }
            //}
            return dt_ret;
        }

        public DataTable GetLabWork(string lab_cd = "")
        {
            string query = string.Format(@"
                SELECT	*
                  FROM	LIS_WS_CD_LIST_V ");
            if (lab_cd != "")
            {
                query += " WHERE WORK_GB='" + lab_cd + "'";
            }
            DataTable dt_ret = new DataTable();

            //using (SqlConnection conn = new SqlConnection(conn_string))
            //{
            //    SqlDataAdapter adpater = new SqlDataAdapter(query, conn);
            //    try
            //    {
            //        adpater.Fill(dt_ret);
            //    }
            //    catch (Exception ex)
            //    {
            //        SysFunc.SaveSystemLog(string.Format("{0}-> {1}", query, ex.Message));
            //    }
            //}
            return dt_ret;
        }


        public DataTable GetLabList()
        {
            string query = string.Format(@"
                SELECT	*
                  FROM	LIS_WORK_LIST_V");
            DataTable dt_ret = new DataTable();

            //using (SqlConnection conn = new SqlConnection(conn_string))
            //{
            //    SqlDataAdapter adpater = new SqlDataAdapter(query, conn);
            //    try
            //    {
            //        adpater.Fill(dt_ret);
            //    }
            //    catch (Exception ex)
            //    {
            //        SysFunc.SaveSystemLog(string.Format("{0}-> {1}", query, ex.Message));
            //    }
            //}
            return dt_ret;
        }


        public MyBact.work_listDataTable Download_Micro(string spc_no)
        {
            MyBact.work_listDataTable table = new MyBact.work_listDataTable();

            return table;

        }

        public MyBact.work_listDataTable Download_Micro_Worklist(DateTime dtp_from, DateTime dtp_to, List<string> exam_list, string equp_cd)
        {
            MyBact.work_listDataTable dt_order = new MyBact.work_listDataTable();

            return dt_order;

        }

        public DataTable GetQC(string spc_no)
        {
            string query = string.Format(@"
                SELECT	DISTINCT QC_GB
                  FROM	LIS_INTERFACE1_V
                WHERE   BCODE_NO = {0}",spc_no);
            DataTable dt_ret = new DataTable();

            //using (SqlConnection conn = new SqlConnection(conn_string))
            //{
            //    SqlDataAdapter adpater = new SqlDataAdapter(query, conn);
            //    try
            //    {
            //        adpater.Fill(dt_ret);
            //    }
            //    catch (Exception ex)
            //    {
            //        SysFunc.SaveSystemLog(string.Format("{0}-> {1}", query, ex.Message));
            //    }
            //}
            return dt_ret;
        }

        public string GetSP10_Value(string spc_no)
        {
            string WBC_Result = "", RBC_Result = "", HCT_Result = "", SP_Result ="";

            //wbc Result
            string query = string.Format(@"
                SELECT	RESULT_NM
                  FROM	LIS_INTERFACE1_V
                WHERE   BCODE_NO = {0} AND ORD_CD = 'LB1050'", spc_no);

            DataTable WBC_TEMP = Select(query);

            if (WBC_TEMP.Rows.Count > 0)
            {
                WBC_Result = WBC_TEMP.Rows[0][0].ToString();
                WBC_Result = WBC_Result.Replace(".", "");
            }

            string query2 = string.Format(@"
                SELECT	RESULT_NM
                  FROM	LIS_INTERFACE1_V
                WHERE   BCODE_NO = {0} AND ORD_CD = 'LB1040'", spc_no);

            DataTable RBC_TEMP = Select(query2);

            if (RBC_TEMP.Rows.Count > 0)
            {
                RBC_Result = RBC_TEMP.Rows[0][0].ToString();
                RBC_Result = RBC_Result.Replace(".", "");
            }

            string query3 = string.Format(@"
                SELECT	RESULT_NM
                  FROM	LIS_INTERFACE1_V
                WHERE   BCODE_NO = {0} AND ORD_CD = 'LB1020'", spc_no);

            DataTable HCT_TEMP = Select(query3);

            if (HCT_TEMP.Rows.Count > 0)
            {
                HCT_Result = HCT_TEMP.Rows[0][0].ToString();
                HCT_Result = HCT_Result.Replace(".", "");
            }

            SP_Result = WBC_Result + "|" + RBC_Result + "|" + HCT_Result;


            return SP_Result;
        }
        public int Upload_MIC_Blood_Rst(
            string info, 
            string spc_no, 
            string exam_cd, 
            string result, 
            string user_id, 
            string work_cd, 
            string host_ip, 
            string equp_cd)
    {
        int n_ret = 0;

        return n_ret;
    }

        public int UpdateMicro(
                    string spc_no, 
                    string exam_cd,
                    string org_seq, 
                    string org_code, 
                    string org_nm, 
                    string anti_code, 
                    string anti_name, 
                    string ris, 
                    string mic, 
                    string equp_cd,
                    string exam_dt, 
                    string user_id)
        {
            int n_ret = 0;

            return n_ret;
        }
        
    }
    
    public class sex_age
    {
        public string sex { get; set; }
        public string age { get; set; }
        public sex_age()
        {
            this.sex = "M";
            this.age = "0";
        }
        public void calc(string jumin)
        {
            string p = "";
            string brith = "";
            if (jumin.Length >= 7)
            {
                p = jumin.Substring(6, 1);
                if (p == "2" || p == "4" || p == "6" || p == "7")
                {
                    this.sex = "M";
                }

                if (p == "1" || p == "2" || p == "5" || p == "6")
                {
                    brith = "19" + jumin.Substring(0, 2) + "-" + jumin.Substring(2, 2) + "-" + jumin.Substring(4, 2);
                }
                else
                {
                    brith = "20" + jumin.Substring(0, 2) + "-" + jumin.Substring(2, 2) + "-" + jumin.Substring(4, 2);
                }
                try
                {
                    this.age = string.Format("{0:0}", (DateTime.Today - DateTime.Parse(brith)).TotalDays / 365);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
        }
    }
}
