CREATE TABLE [tbl_Datasheets] (
  [FileID] AUTOINCREMENT,
  [File_Category] VARCHAR (15),
  [File_Group] VARCHAR (50),
  [File_Code] VARCHAR (50),
  [File_Description] VARCHAR (50),
  [File_Name] VARCHAR (100),
  [File_Path] VARCHAR (255),
  [Sort_Order] LONG ,
  [Inactive] UNSIGNED BYTE 
)
