' 장비ID 찾기 
SELECT Distinct b.AnalyteID, a.lablottestid,c.name,b.MethodID,b.ReagentID, b.UnitID, b.TemperatureID ,b.InstrumentID 
  FROM LabLotTest a, test b, analyte c
 WHERE a.Labid = '506927'
   AND a.Lotid = '6205180'
   AND a.testid = b.testid 
   AND b.AnalyteID = c.AnalyteID 
 ORDER BY a.lablottestid

