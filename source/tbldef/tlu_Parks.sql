CREATE TABLE [tlu_Parks] (
  [ParkCode] VARCHAR (4) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ParkName] VARCHAR (50),
  [ParkState] VARCHAR (2)
)
