IF EXISTS (SELECT name FROM sysobjects WHERE name = 'csp_lip_removetablesandfields' AND UPPER(type) = 'P')
   DROP PROCEDURE [csp_lip_removetablesandfields]
GO

-- Written by: Jonny Springare
-- Created: 2016-03-16

CREATE PROCEDURE [dbo].[csp_lip_removetablesandfields]
	@@idtable INT = NULL
	, @@idfield INT = NULL
	, @@errorMessage NVARCHAR(512) OUTPUT
AS
BEGIN

	-- FLAG_EXTERNALACCESS --
	IF @@idtable IS NOT NULL
	BEGIN
		EXEC lsp_removetable @@idtable = @@idtable
	END
	ELSE IF @@idfield IS NOT NULL
	BEGIN
		EXEC lsp_removefield @@idfield = @@idfield
	END
END

GO

-- Always execute these during installation of procedures to be reached from VBA
EXEC lsp_setdatabasetimestamp
EXEC lsp_refreshldc

GO
