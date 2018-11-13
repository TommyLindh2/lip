IF EXISTS (SELECT name FROM sysobjects WHERE name = 'csp_lip_endinstallation' AND UPPER(type) = 'P')
   DROP PROCEDURE [csp_lip_endinstallation]
GO

-- Written by: Jonny Springare
-- Created: 2016-10-12

-- Modified: 2018-04-09, Fredrik Eriksson: lsp_refreshldc is now always run.

CREATE PROCEDURE [dbo].[csp_lip_endinstallation]
AS
BEGIN
	-- FLAG_EXTERNALACCESS --
	
	IF OBJECT_ID('lsp_setdatabasetimestamp') > 0
	BEGIN
		EXEC lsp_setdatabasetimestamp
	END
	
	EXEC lsp_refreshldc
END

GO

-- Always execute these during installation of procedures to be reached from VBA
EXEC lsp_setdatabasetimestamp
EXEC lsp_refreshldc

GO
