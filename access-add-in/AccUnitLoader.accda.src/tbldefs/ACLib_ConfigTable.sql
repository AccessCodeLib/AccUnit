CREATE TABLE [ACLib_ConfigTable] (
  [PropName] VARCHAR (255) CONSTRAINT [PK_ACLib_ConfigTable] PRIMARY KEY UNIQUE NOT NULL,
  [PropValue] VARCHAR (255),
  [PropRemarks] LONGTEXT 
)
