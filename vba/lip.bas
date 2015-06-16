Attribute VB_Name = "lip"


Private Const BaseURL As String = "http://limebootstrap.lundalogik.com"
Private Const ApiURL As String = "/api/apps/"
Private IndentLenght As String
Private Indent As String

Public Sub Upgrade(Optional PackageName As String, Optional Path As String)
On Error GoTo Errorhandler:
    If PackageName = "" Then
        'Upgrade all packages
        Call InstallFromPackageFile
    Else
        'Upgrade specific package
        Call InstallPackage(PackageName, , True)
    End If
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.Upgrade")
End Sub

Public Sub InstallPackage(Optional PackageName As String, Optional Path As String, Optional Upgrade As Boolean)
On Error GoTo Errorhandler
    Dim Package As Object
    Dim PackageVersion As Double
    
    IndentLenght = "  "
    
    'Check if first use ever
    If Dir(WebFolder + "package.json") = "" Then
        Debug.Print "No package.json found, assuming fresh install"
        Call InstallLIP
    End If
    
    PackageName = LCase(PackageName)
    
    Debug.Print "====== LIP Install: " + PackageName + " ======"
    
    Debug.Print "Looking for package: '" + PackageName + "'"
    Set Package = SearchForPackageOnStores(PackageName, Path)
    If Package Is Nothing Then
        Exit Sub
    Else
        Debug.Print PackageName + " " + Format(PackageVersion, "0.0") + " package found."
        Set Package = Package.Item("info")
    End If
    
    'Parse result from store
    PackageVersion = findNewestVersion(Package.Item("versions"))
    
    'Check of package already exsists
    If Not Upgrade Then
        If CheckForLocalInstalledPackage(PackageName, PackageVersion) = True Then
            Exit Sub
        End If
    End If
    
    'Install dependecies
    
    If Package.Exists("dependencies") Then
        IncreaseIndent
        Call InstallDependencies(Package)
        DecreaseIndent
    End If
   
    Call InstallPackageComponents(PackageName, PackageVersion, Package)
    
    Debug.Print "==================================="
    
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallPackage")
End Sub

Public Sub InstallApp(Optional Path As String)
On Error GoTo Errorhandler
'    IndentLenght = "  "
'
'    'Check if first use ever
'    If Dir(WebFolder + "package.json") = "" Then
'        Debug.Print "No package.json found, assuming fresh install"
'        Call InstallLIP
'    End If
    
    Debug.Print ("Not implemented")
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallApp")
End Sub


Public Sub InstallFromZip(Path As String)
On Error GoTo Errorhandler

'    'TODO:
'    'Check if valid path
'
'    'Check if first use ever
'    If Dir(WebFolder + "package.json") = "" Then
'        Debug.Print "No package.json found, assuming fresh install"
'        Call InstallLIP
'    End If
'
'    PackageName = LCase(PackageName)
'
'    Debug.Print "====== LIP Install: " + PackageName + " ======"
    
    'TODO:
    'Implement installation of zip-fil
    'Copy file to actionpads/apps
    'Unzip file
    'Install dependencies
    'Install package
    
    Debug.Print ("Not implemented")
    
'    Debug.Print "==================================="

Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallFromZip")
End Sub

Public Sub InstallFromPackageFile()
On Error GoTo Errorhandler
    Dim LocalPackages As Object
    Dim LocalPackageName As Variant
    
    Debug.Print "Installing dependecies from package.json file..."
    Set LocalPackages = ReadPackageFile().Item("dependencies")
    If LocalPackages Is Nothing Then
        Exit Sub
    End If
    For Each LocalPackageName In LocalPackages.keys
        Call InstallPackage(CStr(LocalPackageName), , True)
    Next LocalPackageName
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallFromPackageFile")
End Sub


Private Sub InstallPackageComponents(PackageName As String, PackageVersion As Double, Package)
On Error GoTo Errorhandler
    'Download and unzip
    Debug.Print "Downloading '" + PackageName + "' files..."
    Call DownloadFile(PackageName)
    Call Unzip(PackageName)
    Debug.Print "Download complete!"
    
    'Install localizations
    If Package.Item("install").Exists("localize") = True Then
        Debug.Print Indent + "Adding localizations..."
        IncreaseIndent
        Call InstallLocalize(Package.Item("install").Item("localize"))
        DecreaseIndent
          
    End If
    
    'Install VBA
    If Package.Item("install").Exists("vba") = True Then
        Debug.Print Indent + "Adding VBA modules, forms and classes..."
        IncreaseIndent
        Call InstallVBAComponents(PackageName, Package.Item("install").Item("vba"))
        DecreaseIndent
    End If
    
    If Package.Item("install").Exists("tables") = True Then
        IncreaseIndent
        Call InstallFieldsAndTables(Package.Item("install").Item("tables"))
        DecreaseIndent
    End If
    
    If Package.Item("install").Exists("sql") = True Then
        IncreaseIndent
        Call InstallSQL(Package.Item("install").Item("sql"))
        DecreaseIndent
    End If
    'Update package.json
    Call WriteToPackageFile(PackageName, PackageVersion)
    
    Debug.Print Indent + "Installation of " + PackageName + " done!"
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallPackageComponents")
End Sub



Private Sub InstallDependencies(Package As Object)
On Error GoTo Errorhandler
    Dim DependecyName As Variant
    Dim LocalPackage As Object
    Debug.Print Indent + "Dependencies found! Installing..."
    IncreaseIndent
    For Each DepedencyName In Package.Item("dependencies").keys()
        Set LocalPackage = FindPackageLocally(CStr(DepedencyName))
        If LocalPackage Is Nothing Then
            Debug.Print Indent + "Installing dependency: " + CStr(DependecyName)
            Call InstallPackage(CStr(DepedencyName))
        ElseIf Val(LocalPackage.Item(PackageName)) < Val(Package.Item("dependencies").Item(PackageName)) Then
            Call InstallPackage(CStr(DepedencyName), , True)
        Else
        End If
    Next DepedencyName
    DecreaseIndent
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallDependencies")
End Sub

Private Function SearchForPackageOnStores(PackageName As String, Optional Path As String) As Object
On Error GoTo Errorhandler
    Dim sJSON As String
    Dim oJSON As Object
    
    sJSON = getJSON(BaseURL + ApiURL + PackageName + "/")
    Set oJSON = parseJSON(sJSON)
    If Not oJSON.Item("error") = "" Then
        Debug.Print PackageName + " package not found!"
        Set SearchForPackageOnStores = Nothing
        Exit Function
    End If
    If oJSON.Item("info").Item("install") Is Nothing Then
        Debug.Print "Package has no valid install instructions!"
        Set SearchForPackageOnStores = Nothing
        Exit Function
    End If
    Set SearchForPackageOnStores = oJSON
Exit Function
Errorhandler:
    Set SearchForPackageOnStores = Nothing
    Call UI.ShowError("lip.SearchForPackageOnStores")
End Function

Private Function CheckForLocalInstalledPackage(PackageName As String, PackageVersion As Double) As Boolean
On Error GoTo Errorhandler
    Dim LocalPackages As Object
    Dim LocalPackage As Object
    Dim LocalPackageVersion As Double
    Dim LocalPackageName As Variant
    
    Set LocalPackage = FindPackageLocally(PackageName)
        
    If Not LocalPackage Is Nothing Then
        LocalPackageVersion = Val(LocalPackage.Item(PackageName))
        If PackageVersion = LocalPackageVersion Then
            Debug.Print "Current version of" + PackageName + " is already installed, please use lip.upgrade to reinstall package"
            Debug.Print "==================================="
            CheckForLocalInstalledPackage = True
            Exit Function
        ElseIf PackageVersion < LocalPackageVersion Then
            Debug.Print "Package " + PackageName + " is already installed, please use lip.upgrade to upgrade package from " + Format(PackageVersion, "0.0") + " -> " + Format(LocalPackageVersion, "0.0")
            Debug.Print "==================================="
            CheckForLocalInstalledPackage = True
            Exit Function
        Else
            Debug.Print "A newer version of " + PackageName + " is already installed. Remote: " + Format(PackageVersion, "0.0") + " ,Local: " + Format(LocalPackageVersion, "0.0") + " .Please use vip.upgrade to reinstall package"
            Debug.Print "==================================="
            CheckForLocalInstalledPackage = True
            Exit Function
        End If
    End If
    CheckForLocalInstalledPackage = False
Exit Function
Errorhandler:
    Call UI.ShowError("lip.CheckForLocalInstalledPackages")
End Function

Private Function getJSON(sURL As String) As String
On Error GoTo Errorhandler
    Dim qs As String
    qs = CStr(Rnd() * 1000000#)
    Dim oXHTTP As Object
    Dim s As String
    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    oXHTTP.Open "GET", sURL + "?" + qs, False
    oXHTTP.Send
    getJSON = oXHTTP.responseText
Exit Function
Errorhandler:
    getJSON = ""
End Function

Private Function parseJSON(sJSON As String) As Object
On Error GoTo Errorhandler
    Dim oJSON As Object
    Set oJSON = JSON.parse(sJSON)
    Set parseJSON = oJSON
Exit Function
Errorhandler:
    Set parseJSON = Nothing
    Call UI.ShowError("lip.parseJSON")
End Function

Private Function findNewestVersion(oVersions As Object) As Double
On Error GoTo Errorhandler
    Dim NewestVersion As Double
    Dim Version As Object
    NewestVersion = -1
    
    For Each Version In oVersions
        If Val(Version.Item("version")) > NewestVersion Then
            NewestVersion = Val(Version.Item("version"))
        End If
    Next Version
    findNewestVersion = NewestVersion
Exit Function
Errorhandler:
    findNewestVersion = -1
    Call UI.ShowError("lip.findNewestVersion")
End Function

Private Sub InstallLocalize(oJSON As Object)
On Error GoTo Errorhandler
    Dim Localize As Object
        
    For Each Localize In oJSON
        Call AddOrCheckLocalize( _
            Localize.Item("owner"), _
            Localize.Item("context"), _
            "", _
            Localize.Item("en-us"), _
            Localize.Item("sv"), _
            Localize.Item("no"), _
            Localize.Item("fi") _
        )
    Next Localize
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallLocalize")
End Sub

Private Sub InstallSQL(oJSON As Object)
On Error GoTo Errorhandler
    Dim Sql As Object
    Debug.Print Indent + "Installing SQL..."
    IncreaseIndent
    For Each Sql In oJSON
        Debug.Print Indent + "Add: " + Sql.Item("name")
    Next Sql
    DecreaseIndent
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallSQL")
End Sub


Private Sub InstallFieldsAndTables(oJSON As Object)
On Error GoTo Errorhandler
    Dim table As Object
    Dim field As Object
    Dim oClass As LDE.Class
    Debug.Print "Adding fields and tables..."
    IncreaseIndent
    ' Create tables
    For Each table In oJSON
        If Database.Classes.Exists(table.Item("name")) Then
            Debug.Print Indent + "Table '" + table.Item("name") + "' requirement is met"
        Else
            Debug.Print Indent + "Table '" + table.Item("name") + "' needs to be created."
            Call AddTable(table)
        End If
    Next table
    
    ' Create fields
    For Each table In oJSON
        If Database.Classes.Exists(table.Item("name")) Then
            Set oClass = Database.Classes(table.Item("name"))
        End If
        IncreaseIndent
        For Each field In table.Item("fields")
            If oClass Is Nothing Then
                Debug.Print Indent + "Add field: " + field.Item("name")
                Call AddField(table.Item("name"), field)
            ElseIf oClass.Fields.Exists(field.Item("Name")) Then
                Debug.Print Indent + "Field: " + field.Item("name") + " requirement is met"
            Else
                Debug.Print Indent + "Add field: " + field.Item("name")
                Call AddField(table.Item("name"), field)
            End If
        Next field
        DecreaseIndent
    Next table
    DecreaseIndent
Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallFieldsAndTables")
End Sub

Private Sub AddTable(table As Object)
On Error GoTo Errorhandler
    Dim oProc As LDE.Procedure
    Set oProc = Database.Procedures("csp_lip_createtable")
    oProc.Parameters("@@tablename").InputValue = table.Item("name")
    If table.Exists("localname_singular") Then   '##TODO: Rebuild into for each loop
        oProc.Parameters("@@localnamesingularsv").InputValue = table.Item("localname_singular").Item("sv")
        oProc.Parameters("@@localnamesingularenus").InputValue = table.Item("localname_singular").Item("en-us")
    End If
    If table.Exists("localname_plural") Then   '##TODO: Rebuild into for each loop
        oProc.Parameters("@@localnamepluralsv").InputValue = table.Item("localname_plural").Item("sv")
        oProc.Parameters("@@localnamepluralenus").InputValue = table.Item("localname_plural").Item("en-us")
    End If
    
    '##TODO: Loop over "attributes"
    
    Call oProc.Execute(False)
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.AddTable")
End Sub

Private Sub AddField(tableName As String, field As Object)
On Error GoTo Errorhandler
    Dim oProc As LDE.Procedure
    Set oProc = Database.Procedures("csp_lip_createfield")
    oProc.Parameters("@@tablename").InputValue = tableName
    oProc.Parameters("@@fieldname").InputValue = field.Item("name")
    If field.Exists("localname") Then   '##TODO: Rebuild into for each loop
        oProc.Parameters("@@localnamesv").InputValue = field.Item("localname").Item("sv")
        oProc.Parameters("@@localnameenus").InputValue = field.Item("localname").Item("en-us")
    End If
    oProc.Parameters("@@type").InputValue = field.Item("type")
    If field.Exists("attributes") Then
        oProc.Parameters("@@defaultvalue").InputValue = field.Item("attributes").Item("defaultvalue")
    End If
    
    Call oProc.Execute(False)
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.AddField")
End Sub


Private Sub DownloadFile(PackageName As String)
On Error GoTo Errorhandler
    Dim qs As String
    qs = CStr(Rnd() * 1000000#)
    Dim downloadURL As String
    downloadURL = BaseURL + ApiURL + PackageName + "/download/"
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", downloadURL + "?" + qs, False
    WinHttpReq.Send
    
    myURL = WinHttpReq.responseBody
    If WinHttpReq.status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile WebFolder + "apps\" + PackageName + ".zip", 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.DownloadFile")
End Sub

Private Sub Unzip(PackageName)
On Error GoTo Errorhandler
    Dim FSO As Object
    Dim oApp As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String

    Fname = WebFolder + "apps\" + PackageName + ".zip"
    FileNameFolder = WebFolder & "apps\" & PackageName & "\"

    On Error Resume Next
    Set FSO = CreateObject("scripting.filesystemobject")
    'Delete files
    FSO.DeleteFile FileNameFolder & "*.*", True
    'Delete subfolders
    FSO.DeleteFolder FileNameFolder & "*.*", True
    
    'Make the normal folder in DefPath
    MkDir FileNameFolder
    
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).Items
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.Unzip")
End Sub

Private Sub InstallVBAComponents(PackageName As String, VBAModules As Object)
On Error GoTo Errorhandler
    Dim VBAModule As Object
    IncreaseIndent
    For Each VBAModule In VBAModules
        Call addModule(PackageName, VBAModule.Item("name"), VBAModule.Item("relPath"))
        Debug.Print Indent + "Added " + VBAModule.Item("name")
    Next VBAModule
    DecreaseIndent
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallVBAComponents")
End Sub

Private Sub addModule(PackageName As String, ModuleName As String, RelPath As String)
On Error GoTo Errorhandler
    If PackageName <> "" And ModuleName <> "" Then
        Dim VBComps As Object
        Dim Path As String
        
        'Set VBComps = CreateObject("VBIDE.VBComponents")
        'Debug.Print "'Microsoft Visual Basic for Applications Extensibility 5.4' missing. Please add the reference (Tools>References)"
        Set VBComps = Application.VBE.ActiveVBProject.VBComponents
        If ComponentExists(ModuleName, VBComps) = True Then
            VBComps.Item(ModuleName).Name = ModuleName & "OLD"
            Call VBComps.Remove(VBComps.Item(ModuleName & "OLD"))
        End If
        Path = WebFolder + "apps\" + PackageName + "\" + RelPath
     
        Call Application.VBE.ActiveVBProject.VBComponents.Import(Path)
    End If
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.addModule")
End Sub

Private Function ComponentExists(ComponentName As String, VBComps As Object) As Boolean
On Error GoTo Errorhandler
    Dim VBComp As Object
    
    'Set VBComp = CreateObject("VBIDE.VBComponent")
    For Each VBComp In VBComps
        If VBComp.Name = ComponentName Then
             ComponentExists = True
             Exit Function
        End If
    Next VBComp
    
    ComponentExists = False
    
    Exit Function
Errorhandler:
    Call UI.ShowError("lip.ComponentExists")
End Function

Private Sub WriteToPackageFile(PackageName, Version)
On Error GoTo Errorhandler
    Dim oJSON As Object
    Dim Line As Variant
    Set oJSON = ReadPackageFile
    
    oJSON.Item("dependencies").Item(PackageName) = Version

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(WebFolder + "package.json", True)
    For Each Line In Split(PrettyPrintJSON(JSON.toString(oJSON)), vbCrLf)
        a.WriteLine Line
    Next Line
    a.Close
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.WriteToPackageFile")
End Sub

Private Function PrettyPrintJSON(JSON As String) As String
On Error GoTo Errorhandler
    Dim i As Integer
    Dim Indent As String
    Dim PrettyJSON As String
    Dim InsideQuotation As Boolean
    
    For i = 1 To Len(JSON)
        Select Case Mid(JSON, i, 1)
            Case """"
                PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
                If InsideQuotation = False Then
                    InsideQuotation = True
                Else
                    InsideQuotation = False
                End If
            Case "{", "["
                If InsideQuotation = False Then
                    Indent = Indent + "    " ' Add to indentation
                    PrettyJSON = PrettyJSON + "{" + vbCrLf + Indent
                Else
                    PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
                End If
            Case "}", "["
                If InsideQuotation = False Then
                    Indent = Left(Indent, Len(Indent) - 4) 'Remove indentation
                    PrettyJSON = PrettyJSON + vbCrLf + Indent + "}"
                Else
                    PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
                End If
            Case ","
                If InsideQuotation = False Then
                    PrettyJSON = PrettyJSON + "}" + vbCrLf + Indent
                Else
                    PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
                End If
            Case Else
                PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
        End Select
    Next i
    PrettyPrintJSON = PrettyJSON
    
    Exit Function
Errorhandler:
    PrettyPrintJSON = ""
    Call UI.ShowError("lip.PrettyPrintJSON")
End Function

Private Function ReadPackageFile() As Object
On Error GoTo Errorhandler
    Dim sJSON As String
    Dim oJSON As Object
    sJSON = getJSON(WebFolder + "package.json") '"package.json")
    
    If sJSON = "" Then
        Debug.Print "Error: No package.json found!"
        Set ReadPackageFile = Nothing
        Exit Function
    End If
    
    Set oJSON = JSON.parse(sJSON)
    Set ReadPackageFile = oJSON
    
    Exit Function
Errorhandler:
    Set ReadPackageFile = Nothing
    Call UI.ShowError("lip.ReadPackageFile")
End Function

Private Function FindPackageLocally(PackageName As String) As Object
On Error GoTo Errorhandler
    Dim InstalledPackages As Object
    Dim Package As Object
    Dim ReturnDict As New Scripting.Dictionary
    Set InstalledPackages = ReadPackageFile.Item("dependencies")
        If InstalledPackages.Exists(PackageName) = True Then
            Call ReturnDict.Add(PackageName, InstalledPackages.Item(PackageName))
            Set FindPackageLocally = ReturnDict
            Exit Function
        End If
    Set FindPackageLocally = Nothing
    Exit Function
Errorhandler:
    Set FindPackageLocally = Nothing
    Call UI.ShowError("lip.FindPackageLocally")
End Function

Private Sub CreateANewPackageFile()
On Error GoTo Errorhandler
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(WebFolder + "package.json", True)
    a.WriteLine "{"
    a.WriteLine "    ""dependencies"":{}"
    a.WriteLine "}"
    a.Close
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.CreateNewPackageFile")
End Sub

Private Sub InstallLIP()
On Error GoTo Errorhandler
    Debug.Print "Installing JSON-lib..."
    Call DownloadFile("vba_json")
    Call Unzip("vba_json")
    Call addModule("vba_json", "JSON", "JSON.bas")
    Call addModule("vba_json", "cStringBuilder", "cStringBuilder.cls")
    
    Debug.Print "Creating a new package.json file..."
    Call CreateANewPackageFile
    Call WriteToPackageFile("vba_json", 1)

    Debug.Print "Install of LIP complete!"
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.InstallLIP")
End Sub

Private Function AddOrCheckLocalize(sOwner As String, sCode As String, sDescription As String, sEN_US As String, sSV As String, sNO As String, sFI As String) As Boolean
On Error GoTo Errorhandler
    Dim oFilter As New LDE.Filter
    Dim oRecs As New LDE.Records
    
    Call oFilter.AddCondition("owner", lkOpEqual, sOwner)
    Call oFilter.AddCondition("code", lkOpEqual, sCode)
    oFilter.AddOperator lkOpAnd
    
    If oFilter.HitCount(Database.Classes("localize")) = 0 Then
        Debug.Print (Indent + "Localization " & sOwner & "." & sCode & " not found, creating new!")
        Dim oRec As New LDE.Record
        Call oRec.Open(Database.Classes("localize"))
        oRec.Value("owner") = sOwner
        oRec.Value("code") = sCode
        oRec.Value("context") = sDescription
        oRec.Value("sv") = sSV
        oRec.Value("en-us") = sEN_US
        oRec.Value("no") = sNO
        oRec.Value("fi") = sFI
        Call oRec.Update
    ElseIf oFilter.HitCount(Database.Classes("localize")) = 1 Then
    Debug.Print (Indent + "Updating localization " & sOwner & "." & sCode)
        Call oRecs.Open(Database.Classes("localize"), oFilter)
        oRecs(1).Value("owner") = sOwner
        oRecs(1).Value("code") = sCode
        oRecs(1).Value("context") = sDescription
        oRecs(1).Value("sv") = sSV
        oRecs(1).Value("en-us") = sEN_US
        oRecs(1).Value("no") = sNO
        oRecs(1).Value("fi") = sFI
        Call oRecs.Update
        
    Else
        Call MsgBox("There are multiple copies of " & sOwner & "." & sCode & "  which is bad! Fix it", vbCritical, "To many translations makes Jack a dull boy")
    End If
    
    Set Localize.dicLookup = Nothing
    AddOrCheckLocalize = True
    Exit Function
Errorhandler:
    Debug.Print ("Error while validating or adding Localize")
    AddOrCheckLocalize = False
End Function

Private Sub IncreaseIndent()
On Error GoTo Errorhandler
    Indent = Indent + IndentLenght
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.IncreaseIndent")
End Sub

Private Sub DecreaseIndent()
On Error GoTo Errorhandler
    Indent = Left(Indent, Len(Indent) - Len(IndentLenght))
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.DecreaseIndent")
End Sub


