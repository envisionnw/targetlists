Operation =6
Option =0
Begin InputTables
    Name ="qryAnnualCompleteTgtSpeciesListsAllParks"
End
Begin OutputColumns
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.Species_Name"
    GroupLevel =2
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.ARCHPriority"
    GroupLevel =1
    Alias ="CountOfFamily"
    Expression ="Count(qryAnnualCompleteTgtSpeciesListsAllParks.Family)"
    Alias ="Total Of Family"
    Expression ="Count(qryAnnualCompleteTgtSpeciesListsAllParks.Family)"
    GroupLevel =2
End
Begin Groups
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.Species_Name"
    GroupLevel =2
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.Family"
    GroupLevel =2
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.utah_species"
    GroupLevel =2
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.Co_Species"
    GroupLevel =2
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.Wy_Species"
    GroupLevel =2
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.Master_Common_Name"
    GroupLevel =2
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.Master_Plant_Code_FK"
    GroupLevel =2
    Expression ="qryAnnualCompleteTgtSpeciesListsAllParks.ARCHPriority"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x326fbcddd91c4a4cb022613a36bfcb64
End
Begin
    Begin
        dbText "Name" ="[Species_Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Total Of Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd756556dda8f044ba1272d2f194dfbb5
        End
    End
    Begin
        dbText "Name" ="X"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfFamily"
        dbBinary "GUID" = Begin
            0xe6c6319530721240b0df558ed7794cab
        End
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.[Species_Name]"
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
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.ARCHPriority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.[ARCHPriority]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesListsAllParks.[BLCAPriority]"
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
    Right =1679
    Bottom =324
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="qryAnnualCompleteTgtSpeciesListsAllParks"
        Name =""
    End
End
