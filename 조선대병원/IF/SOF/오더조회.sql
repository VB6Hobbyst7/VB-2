 select M.PART_JUBSU_DATE                       
      , M.PART_JUBSU_TIME                       
      , M.BUNHO                                 
      , M.SUNAME                                
      , M.AGE                                   
      , M.SEX                                   
      , M.SPECIMEN_CODE                         
      , M.GWA_NAME                              
      , R.HANGMOG_CODE                          
      , E.GUMSA_NAME                            
      , E.JANGBI_OUT_CODE                       
      , E.JANGBI_CODE                           
      , R.LAB_NO                                
      , R.CONFIRM_YN                            
      , R.CPL_RESULT                            
      , R.JANGBI_YN                             
      , R.JANGBI_CODE                           
 from MEDI.CPL3020 R                                 
    , MEDI.CPL2010 M                                 
    , MEDI.CPL0101 E                                 
 where R.SPECIMEN_SER ='0000000000' 
   and NVL(R.CONFIRM_YN, 'N') = 'N'         
   and R.JANGBI_CODE = 'IK'   
   and E.JANGBI_OUT_CODE is Not Null            
   and R.SPECIMEN_SER = M.SPECIMEN_SER          
   and R.SPECIMEN_CODE = M.SPECIMEN_CODE        
   and R.HANGMOG_CODE = E.HANGMOG_CODE          
