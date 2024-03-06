CREATE TABLE [USys_AppFiles] (
  [id] VARCHAR (50),
  [BitInfo] VARCHAR (255),
  [version] VARCHAR (20),
  [file] LONGBINARY ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([id], [BitInfo])
)
