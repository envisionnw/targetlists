﻿dbMemo "SQL" ="SELECT tbl_Target_Species.Park_Code AS Park, tbl_Target_Species.Target_Year AS T"
    "gtYear, Master_Plant_Code_FK, Species_Name, Priority, Transect_Only, Target_Area"
    "_ID\015\012FROM tbl_Target_Species\015\012WHERE (((tbl_Target_Species.Target_Yea"
    "r) = CInt(2013)) And ((LCase([tbl_Target_Species].[Park_Code])) = LCase('BRCA'))"
    ")\015\012ORDER BY tbl_Target_Species.Species_Name;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
