
IF EXISTS (SELECT name FROM sysobjects WHERE name = 'csp_lip_addrelations' AND UPPER(type) = 'P')
   DROP PROCEDURE [csp_lip_addrelations]
GO

-- Written by: Jonny Springare
-- Created: 2015-12-18

-- Creates a relation between two fields. The two fields must have been created during this particular installation.

-- Modified: 2018-03-12: Changed from error to warning when relation already existed.

CREATE PROCEDURE [dbo].[csp_lip_addrelations]
	@@table1 NVARCHAR(64)
	, @@field1 NVARCHAR(64) = NULL
	, @@table2 NVARCHAR(64)
	, @@field2 NVARCHAR(64) = NULL
	, @@createdfields NVARCHAR(MAX)
	, @@errorMessage NVARCHAR(MAX) OUTPUT
	, @@warningMessage NVARCHAR(MAX) OUTPUT
AS
BEGIN

	-- FLAG_EXTERNALACCESS --

	DECLARE @idfield1 INT
	DECLARE @idtable1 INT
	DECLARE @fieldtype1 INT
	
	DECLARE @idfield2 INT
	DECLARE @idtable2 INT
	DECLARE @fieldtype2 INT
	
	DECLARE @linebreak NVARCHAR(2)
	SET @linebreak = CHAR(13) + CHAR(10)
	SET @@errorMessage = N''
	SET @@warningMessage = N''
	
	--Get id for fieldtype relation
	DECLARE @fieldtypeRelation INT
	SELECT @fieldtypeRelation = idfieldtype
	FROM fieldtype
	WHERE [name] = N'relation'
		AND [active] = 1
		AND [creatable] = 1
	
	--Get id:s
	EXEC lsp_getfield @@idfield=@idfield1 OUTPUT, @@name=@@field1, @@table=@@table1, @@fieldtype=@fieldtype1 OUTPUT
	EXEC lsp_gettable @@idtable=@idtable1 OUTPUT, @@name=@@table1
	
	EXEC lsp_getfield @@idfield=@idfield2 OUTPUT, @@name=@@field2, @@table=@@table2, @@fieldtype=@fieldtype2 OUTPUT
	EXEC lsp_gettable @@idtable=@idtable2 OUTPUT, @@name=@@table2
	

	-- PART 1: Validate and check stuff --
	--------------------------------------

	--Check if fields exist
	IF @idfield1 IS NULL
	BEGIN
		SET @@errorMessage = N'ERROR: ' + @@table1 + '.' + @@field1 + N' has not been created during this installation, relation to ' + @@table2 + '.' + @@field2 + N' cannot be created.'
		RETURN
	END

	IF @idfield2 IS NULL
	BEGIN
		SET @@errorMessage = N'ERROR: ' + @@table2 + '.' + @@field2 + N' has not been created during this installation, relation to ' + @@table1 + '.' + @@field1 + N' cannot be created.'
		RETURN
	END


	-- Check if any of the fields existed before this installation
	SET @@createdfields = N';' + @@createdfields
	
	-- Check if field1 existed but not field2
	IF CHARINDEX(N';' + CONVERT(NVARCHAR(MAX), @idfield1) + N';', @@createdfields) > 0 AND CHARINDEX(N';' + CONVERT(NVARCHAR(MAX), @idfield2) + N';', @@createdfields) <= 0
	BEGIN
		SET @@errorMessage = N'ERROR: Cannot create relation between ' + @@table1 + '.' + @@field1 + N' and ' + @@table2 + '.' + @@field2 + N', since ' + @@table1 + '.' + @@field1 + N' already existed before this installation but ' + @@table2 + '.' + @@field2 + N' did not.'
		RETURN
	END

	-- Check if field2 existed but not field1
	IF CHARINDEX(N';' + CONVERT(NVARCHAR(MAX), @idfield1) + N';', @@createdfields) <= 0 AND CHARINDEX(N';' + CONVERT(NVARCHAR(MAX), @idfield2) + N';', @@createdfields) > 0
	BEGIN
		SET @@errorMessage = N'ERROR: Cannot create relation between ' + @@table1 + '.' + @@field1 + N' and ' + @@table2 + '.' + @@field2 + N', since ' + @@table2 + '.' + @@field2 + N' already existed before this installation but ' + @@table1 + '.' + @@field1 + N' did not.'
		RETURN
	END

	-- Check if the fields are relation fields
	IF @fieldtype1 <> @fieldtypeRelation
	BEGIN
		SET @@errorMessage = N'ERROR: ' + @@table1 + '.' + @@field1 + N' is not a relation field/tab, relation to ' + @@table2 + '.' + @@field2 + N' cannot be created.'
		RETURN
	END

	IF @fieldtype2 <> @fieldtypeRelation
	BEGIN
		SET @@errorMessage = N'ERROR: ' + @@table2 + '.' + @@field2 + N' is not a relation field/tab, relation to ' + @@table1 + '.' + @@field1 + N' cannot be created.'
		RETURN
	END

	-- Check if the fields exist in table relationfield
	IF NOT EXISTS (SELECT idrelationfield FROM relationfield WHERE idfield = @idfield1)
	BEGIN
		SET @@errorMessage = N'ERROR: ' + @@table1 + '.' + @@field1 + N' does not exist in table ''relationfield'', relation to ' + @@table2 + '.' + @@field2 + N' cannot be created.'
		RETURN
	END

	IF NOT EXISTS (SELECT idrelationfield FROM relationfield WHERE idfield = @idfield2)
	BEGIN
		SET @@errorMessage = N'ERROR: ' + @@table2 + '.' + @@field2 + N' does not exist in table ''relationfield'', relation to ' + @@table1 + '.' + @@field1 + N' cannot be created.'
		RETURN
	END

	-- Check if one of the two fields already is in a relation (for new relations, there will be a row in relationfieldview but relatedidfield will be NULL).
	DECLARE @relatedidfield INT
	SELECT @relatedidfield = relatedidfield FROM relationfieldview WHERE idfield = @idfield1
	IF @relatedidfield IS NOT NULL
	BEGIN
		IF @relatedidfield = @idfield2
		BEGIN
			SET @@warningMessage = N'Warning: Relation between ' + @@table1 + '.' + @@field1 + ' and ' + @@table2 + '.' + @@field2 + N' already exists.'
		END
		ELSE
		BEGIN
			SET @@errorMessage = N'ERROR: ' + @@table1 + '.' + @@field1 + N' is already in another relation, relation to ' + @@table2 + '.' + @@field2 + N' cannot be created.'
		END
		RETURN
	END

	SET @relatedidfield = NULL		-- Must reset it first, if the below SELECT statement does not return any hits at all, @relatedidfield will remain unchanged. 
	SELECT @relatedidfield = relatedidfield FROM relationfieldview WHERE idfield = @idfield2
	IF @relatedidfield IS NOT NULL
	BEGIN
		SET @@errorMessage = N'ERROR: ' + @@table2 + '.' + @@field2 + N' is already in another relation, relation to ' + @@table1 + '.' + @@field1 + N' cannot be created.'
		RETURN
	END
	--------------------------------------

	-- PART 2: Add the relation --
	------------------------------

	IF @@errorMessage = N'' AND @@warningMessage = N''
	BEGIN
		-- If we arrive here, all is well: Add the relation!
		DECLARE	@return_value INT
		EXEC @return_value = lsp_addrelation
				@@idfield1 = @idfield1,
				@@idtable1 = @idtable1,
				@@idfield2 = @idfield2,
				@@idtable2 = @idtable2
	
		IF @return_value <> 0
		BEGIN
			SET @@errorMessage = N'ERROR: could not create relation between ' + @@table2 + '.' + @@field2 + N' and ' + @@table1 + '.' + @@field1 + N'.'
		END
	END
	------------------------------
END

GO

-- Always execute these during installation of procedures to be reached from VBA
EXEC lsp_setdatabasetimestamp
EXEC lsp_refreshldc

GO
