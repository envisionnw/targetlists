﻿dbMemo "SQL" ="SELECT DISTINCT qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK, qryAnnual"
    "CompleteTgtSpeciesLists.Family, qryAnnualCompleteTgtSpeciesLists.Species_Name, q"
    "ryAnnualCompleteTgtSpeciesLists.utah_species, qryAnnualCompleteTgtSpeciesLists.C"
    "o_Species, qryAnnualCompleteTgtSpeciesLists.Wy_Species, qryAnnualCompleteTgtSpec"
    "iesLists.Master_Common_Name, (qryAnnualCompleteTgtSpeciesLists.Park + \"-\" + qr"
    "yAnnualCompleteTgtSpeciesLists.PriorityTarget) AS ParkPriority, ConcatRelated(\""
    "PriorityTarget\",\"qryAnnualCompleteTgtSpeciesLists\",\"Species_Name=\"\"\" & [S"
    "pecies_Name] & \"\"\" and Park=\"\"\" & [Park] & \"\"\"\",'',\"|\") AS ParkPrior"
    "ities, ConcatRelated(\"ParkPriority\",\"qryAnnualCompleteTgtSpeciesLists\",\"Spe"
    "cies_Name\",'',\"|\") AS ParkPriorities2\015\012FROM qryAnnualCompleteTgtSpecies"
    "Lists;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x168d07f52a87864bb9feae2549a3b03d
End
Begin
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkPriorities"
        dbInteger "ColumnWidth" ="4440"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd9889c2e6375c048961504abd97141f5
        End
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkPriority"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xea22beb07666f14cad13ca4b700dc197
        End
    End
    Begin
        dbText "Name" ="ParkPriorities2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9acba679b02a3c41b15a9d6f37c0cdc9
        End
    End
End
