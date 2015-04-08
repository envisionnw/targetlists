Operation =1
Option =2
Begin InputTables
    Name ="qryAnnualCompleteTgtSpeciesLists"
End
Begin OutputColumns
    Expression ="qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Family"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Species_Name"
    Expression ="qryAnnualCompleteTgtSpeciesLists.utah_species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Co_Species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Wy_Species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Master_Common_Name"
    Alias ="ParkPriorities"
    Expression ="ConcatRelated(\"PriorityTarget\",\"qryAnnualCompleteTgtSpeciesLists\",\"Species_"
        "Name\",'',\"|\")"
End
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
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Family"
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
End
Begin
    State =0
    Left =0
    Top =0
    Right =1699
    Bottom =805
    Left =-1
    Top =-1
    Right =1301
    Bottom =477
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =354
        Bottom =399
        Top =0
        Name ="qryAnnualCompleteTgtSpeciesLists"
        Name =""
    End
End
