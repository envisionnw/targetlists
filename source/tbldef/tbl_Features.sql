CREATE TABLE [tbl_Features] (
  [Feature_PK] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Loc_ID_FK] LONG ,
  [Feature_ID] VARCHAR (1),
  [Feature_description] VARCHAR (255),
  [Feature_directions] LONGTEXT 
)
