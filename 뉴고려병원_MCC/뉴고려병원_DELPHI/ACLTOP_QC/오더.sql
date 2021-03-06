 SELECT DISTINCT  READING_YMD AS HOSPDATE 
                , BCODE_NO AS BARCODE     
                , PTNT_NO AS PID          
                , PTNT_NM AS PNAME        
                , AGE AS AGE              
                , SEX AS SEX              
                , IO_GB AS INOUT          
                , ORD_CD AS ITEM          
                , SP_CD AS SPCCD          
 FROM LIS_INTERFACE1_V                    
 WHERE BCODE_NO = '1234567890' 
   AND STS_CD = '0'                     
 ORDER BY ORD_CD 
