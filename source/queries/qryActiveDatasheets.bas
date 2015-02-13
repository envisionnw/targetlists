dbMemo "SQL" ="SELECT Left(tbl_Datasheets.File_Name,4) AS Park, tbl_Datasheets.File_Code, tbl_D"
    "atasheets.File_Name AS Datasheet, tbl_Datasheets.File_Description, tbl_Datasheet"
    "s.Sort_Order\015\012FROM tbl_Datasheets\015\012WHERE (((tbl_Datasheets.Inactive)"
    "=0))\015\012ORDER BY tbl_Datasheets.Sort_Order;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa5e19d086f23c846b54fbc8ad38753ad
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Datasheets.File_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3463632dc427054999643de3abc13059
        End
    End
    Begin
        dbText "Name" ="tbl_Datasheets.File_Description"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x76be796313b56e4d9dcb036576c6cbce
        End
    End
    Begin
        dbText "Name" ="Datasheet"
        dbInteger "ColumnWidth" ="6720"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe60331dc5c7b654e9763a33890d5f864
        End
    End
    Begin
        dbText "Name" ="Park"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x56f8e18c225ad04fab05b9c26bba1708
        End
    End
    Begin
        dbText "Name" ="tbl_Datasheets.Sort_Order"
        dbLong "AggregateType" ="-1"
    End
End
