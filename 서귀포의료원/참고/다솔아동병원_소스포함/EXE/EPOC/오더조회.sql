 SELECT H141_EXAMPLCE 검사파트     
      , H141_TSAMPLENO 샘플번호    
 　　 , H141_TAKEDAT 수거일자      
      , H141_TAKETM 수거시간       
      , H141_SEQNO 고유번호        
      , H141_TAKESEQ 수거순번      
      , H141_CHARTNO 차트번호      
      , FN_PATIENT_INFO(H141_CHARTNO) 환자성명              
      , FN_PATIENT_INFO(H141_CHARTNO, 'B') 생년월일       
      , FN_SEXAGE(H141_CHARTNO) 성별나이                    
      , H141_VISTDAT 방문일자                               
      , H141_ODRDAT 처방일자                                
      , H141_ODRNO 처방번호                                 
      , H141_ODRSEQ 처방서브번호                            
      , H141_SUGACD 처방코드                                
      , FN_SUGAMST_INFO( H141_SUGACD, 'H') 한글명         
      , FN_SUGAMST_INFO( H141_SUGACD, 'E') 영문명         
      , H141_RSLTYN 결과유무                                
      , H141_NOTYYN 통보유무                                
      , H141_SPECCD                                         
 FROM TB_H141_LISTAKEBODY                                   
    , TB_H131_SPPRESULT                                     
 WHERE H141_TSAMPLENO = '2015009262'            
   AND H141_TAKEDAT between '20151004'                          And '20151005'    
   AND NVL(H141_RSLTYN,' ') IN('N', 'T')                         
   AND H141_SEQNO = H131_SEQNO                                         
   AND (TRIM(H131_RESULT) IS NULL OR TRIM(H131_RESULT) = '결과대기') 
   AND H131_SPPTYPE = 'L010'                                         
