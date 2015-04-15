dbMemo "SQL" ="SELECT DISTINCT tbl_Target_Species.Species_Name, qryAnnualCompleteTgtSpeciesList"
    "sAllParks.Master_Plant_Code_FK, qryAnnualCompleteTgtSpeciesListsAllParks.Species"
    "_Name, qryAnnualCompleteTgtSpeciesListsAllParks.utah_species, qryAnnualCompleteT"
    "gtSpeciesListsAllParks.Co_Species, qryAnnualCompleteTgtSpeciesListsAllParks.Wy_S"
    "pecies, qryAnnualCompleteTgtSpeciesListsAllParks.Master_Common_Name, (qryAnnualC"
    "ompleteTgtSpeciesListsAllParks.ARCHPriority + \"|\" + qryAnnualCompleteTgtSpecie"
    "sListsAllParks.BLCAPriority\015\012+  \"|\" + qryAnnualCompleteTgtSpeciesListsAl"
    "lParks.BRCAPriority + \"|\" + qryAnnualCompleteTgtSpeciesListsAllParks.CANYPrior"
    "ity\015\012+  \"|\" + qryAnnualCompleteTgtSpeciesListsAllParks.CAREPriority + \""
    "|\" + qryAnnualCompleteTgtSpeciesListsAllParks.CEBRPriority\015\012+  \"|\" + qr"
    "yAnnualCompleteTgtSpeciesListsAllParks.COLMPriority + \"|\" + qryAnnualCompleteT"
    "gtSpeciesListsAllParks.CUREPriority\015\012+  \"|\" + qryAnnualCompleteTgtSpecie"
    "sListsAllParks.DINOPriority + \"|\" + qryAnnualCompleteTgtSpeciesListsAllParks.F"
    "OBUPriority\015\012+  \"|\" + qryAnnualCompleteTgtSpeciesListsAllParks.GOSPPrior"
    "ity + \"|\" + qryAnnualCompleteTgtSpeciesListsAllParks.HOVEPriority\015\012+  \""
    "|\" + qryAnnualCompleteTgtSpeciesListsAllParks.NABRPriority + \"|\" + qryAnnualC"
    "ompleteTgtSpeciesListsAllParks.PISPPriority\015\012+  \"|\" + qryAnnualCompleteT"
    "gtSpeciesListsAllParks.TICAPriority + \"|\" + qryAnnualCompleteTgtSpeciesListsAl"
    "lParks.ZIONPriority\015\012) AS ParkPriority\015\012FROM tbl_Target_Species INNE"
    "R JOIN qryAnnualCompleteTgtSpeciesListsAllParks ON tbl_Target_Species.Master_Pla"
    "nt_Code_FK = qryAnnualCompleteTgtSpeciesListsAllParks.Master_Plant_Code_FK;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x08301b61e0508f4f86fe8dbd597d7acf
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkPriority"
        dbInteger "ColumnWidth" ="2940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x91e7a516980bf547bf31b00b68c2dd46
        End
    End
End
