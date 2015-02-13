dbMemo "SQL" ="SELECT Switch(tlu_NCPN_Plants.LU_Code Is Null,\" \",tlu_NCPN_Plants.LU_Code<>\"\""
    ",tlu_NCPN_Plants.LU_Code) AS Code, tlu_NCPN_Plants.Master_Species AS Species, tl"
    "u_NCPN_Plants.Master_PLANT_Code\015\012FROM tlu_NCPN_Plants;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7181a99fdd61ee499e91e1272d2ff5d0
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbd93ec5fffe281499ceb78908087bc23
        End
        dbInteger "ColumnWidth" ="1704"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xabe9b57fcb7ef64598b255922b8dac2f
        End
    End
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x66c6f8479133944db940aec66aa2e039
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_PLANT_Code"
        dbLong "AggregateType" ="-1"
    End
End
