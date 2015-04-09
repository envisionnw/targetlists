dbMemo "SQL" ="SELECT DISTINCT qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK, qryAnnual"
    "CompleteTgtSpeciesLists.LU_Code, qryAnnualCompleteTgtSpeciesLists.Family, qryAnn"
    "ualCompleteTgtSpeciesLists.Species_Name, qryAnnualCompleteTgtSpeciesLists.utah_s"
    "pecies, qryAnnualCompleteTgtSpeciesLists.Co_Species, qryAnnualCompleteTgtSpecies"
    "Lists.Wy_Species, qryAnnualCompleteTgtSpeciesLists.Master_Common_Name, ConcatRel"
    "ated(\"ParkPriority\",\"qryAnnualCompleteTgtSpeciesLists\",\"Species_Name='\"+Sp"
    "ecies_Name+\"'\",'',\"|\") AS ParkPriorities\015\012FROM qryAnnualCompleteTgtSpe"
    "ciesLists;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x67580735b574924785e273813ad36afb
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
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
        dbInteger "ColumnWidth" ="5625"
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
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.LU_Code"
        dbLong "AggregateType" ="-1"
    End
End
