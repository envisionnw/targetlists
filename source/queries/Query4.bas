Operation =1
Option =2
Begin InputTables
    Name ="tbl_Target_Species"
    Name ="tbl_Target_Areas"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Alias ="Park"
    Expression ="tbl_Target_Species.Park_Code"
    Alias ="TgtYear"
    Expression ="tbl_Target_Species.Target_Year"
    Expression ="tbl_Target_Species.Master_Plant_Code_FK"
    Expression ="tbl_Target_Species.Species_Name"
    Expression ="tbl_Target_Species.Priority"
    Expression ="tbl_Target_Species.Transect_Only"
    Expression ="tbl_Target_Species.Target_Area_ID"
    Alias ="Tgt_Area"
    Expression ="tbl_Target_Areas.Target_Area"
    Alias ="Family"
    Expression ="tlu_NCPN_Plants.Master_Family"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Expression ="tlu_NCPN_Plants.utah_species"
    Expression ="tlu_NCPN_Plants.Co_Species"
    Expression ="tlu_NCPN_Plants.Wy_Species"
    Alias ="PriorityTarget"
    Expression ="IIf(tbl_Target_Species.Target_Area_ID>0,tbl_Target_Areas.Target_Area,IIf(tbl_Tar"
        "get_Species.Transect_Only>0,\"Transect\",tbl_Target_Species.Priority))"
    Alias ="ParkPriority"
    Expression ="(tbl_Target_Species.Park_Code+\"-\"+PriorityTarget)"
    Alias ="ParkPriorities"
    Expression ="ConcatRelated(\"ParkPriority\",\"qryAnnualCompleteTgtSpeciesLists\",\"Species_Na"
        "me='\"+Species_Name+\"'\",'',\"|\")"
End
Begin Joins
    LeftTable ="tbl_Target_Species"
    RightTable ="tbl_Target_Areas"
    Expression ="tbl_Target_Species.Target_Area_ID = tbl_Target_Areas.Target_Area_ID"
    Flag =2
    LeftTable ="tbl_Target_Species"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_Target_Species.Master_Plant_Code_FK = tlu_NCPN_Plants.LU_Code"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Target_Species.Species_Name"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x19d197985a33c041b57aeaf0580ad701
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Park"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf219e27b13528d4dadbc2cc40e167601
        End
    End
    Begin
        dbText "Name" ="TgtYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa0046641a09dbb4c93902bc14ca3b338
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Transect_Only"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Target_Area_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tgt_Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x230c8a0ee16f1f4692af28ee4e703ae7
        End
    End
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ec37b94c8738949bf0acb32ce65bd76
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PriorityTarget"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8044312a1365247acdbc98c3db57d9e
        End
    End
    Begin
        dbText "Name" ="ParkPriority"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x820e3fc05454f24d9f8503c794b3ac29
        End
    End
    Begin
        dbText "Name" ="ParkPriorities2"
        dbInteger "ColumnWidth" ="15195"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x07f557305f8de14c8b157454532e7658
        End
    End
    Begin
        dbText "Name" ="ParkPriorities"
        dbInteger "ColumnWidth" ="6135"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0aac388512b32648b956da6868d2f571
        End
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =789
    Bottom =801
    Left =-1
    Top =-1
    Right =773
    Bottom =539
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Target_Species"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Target_Areas"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
