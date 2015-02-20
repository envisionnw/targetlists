dbMemo "SQL" ="SELECT LU_Code, Master_Species, Utah_Species, CO_Species, WY_Species, Master_Com"
    "mon_Name\015\012FROM tlu_NCPN_Plants\015\012WHERE WY_Species LIKE '*pop*';\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x32a45e1457b7ac4c895583a6e4485771
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Utah_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CO_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WY_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
End
