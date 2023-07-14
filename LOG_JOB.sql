
/****** Object:  Table [dbo].[LOG_JOB]    Script Date: 13/07/2023 22:56:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[LOG_JOB](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[TANGGAL_DB] [date] NULL,
	[TGL_TARIK] [datetime] NULL,
	[REGION] [nchar](10) NULL,
	[NAMA_FILE] [text] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
