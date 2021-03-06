Attribute VB_Name = "lip"
Option Explicit

'Lime Package Store, DO NOT CHANGE, used to download system files for LIP
'Please add your own stores in packages.json
Private Const BaseURL As String = "http://api.lime-bootstrap.com"
Private Const PackageStoreApiURL As String = "/addons/"
Private Const AppStoreApiURL As String = "/addons/"

Private Const DefaultInstallPath As String = "packages\"

' Used for setting the installed LIP version in the packages.json file that LIP creates upon first install.
Private Const m_sLIPVersion As String = "1.3.1"


Private IndentLenght As String
Private Indent As String
Private sLog As String

Private m_frmProgress As FormProgress
Private m_progressDouble As Double
Private Const ProgressBarIncrease As Double = (100 / 11)


Public Sub UpgradePackage(Optional PackageName As String, Optional Path As String)
On Error GoTo ErrorHandler:
    If PackageName = "" Then
        'Upgrade all packages
        Call InstallFromPackageFile
    Else
        'Upgrade specific package
        Call Install(PackageName, True)
    End If
Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.UpgradePackage")
End Sub

'Install package/app. Selects packagestore from packages.json
Public Sub Install(PackageName As String, Optional upgrade As Boolean, Optional Simulate As Boolean = True)
    On Error GoTo ErrorHandler
    
    Dim Package As Object
    Dim PackageVersion As Double
    Dim downloadURL As String
    Dim sInstallPath As String
    Dim bOK As Boolean
    Dim bLocalPackage As Boolean
    Dim tempProgress As Double
    Dim tempCaption As String
    
    If Not VerifyLIPInstallation Then
        Call Lime.MessageBox("Verification of LIP failed, installation aborted.", vbCritical)
        Exit Sub
    End If
    
    If m_frmProgress Is Nothing Then
        Set m_frmProgress = New FormProgress
        If Simulate Then
            m_frmProgress.Caption = "Simulating " & PackageName
        Else
            m_frmProgress.Caption = "Installing " & PackageName
        End If
        m_progressDouble = 0
    End If

    Indent = ""
    IndentLenght = "  "
    sLog = ""
    bOK = True
    
    Application.MousePointer = 11

    Call m_frmProgress.show
    
    'Check if first use ever
    If VBA.Dir(ThisApplication.WebFolder + "packages.json") = "" Then
        sLog = sLog + Indent + "No packages.json found, assuming fresh install" + VBA.vbNewLine
        
        tempProgress = m_progressDouble
        tempCaption = m_frmProgress.Caption
        
        m_progressDouble = 0
        Call InstallLIP
        
        If m_frmProgress Is Nothing Then
            Set m_frmProgress = New FormProgress
            m_frmProgress.show
        End If
        m_progressDouble = tempProgress
        m_frmProgress.Caption = tempCaption
    Else
        '##TODO Check if LIP has a new version
        tempCaption = m_frmProgress.Caption
        m_frmProgress.Caption = "Updating LIP"
        Call updateProgressBar("Updating LIP if necessary", m_progressDouble)
        Call UpdateLIPOnNewVersion
        m_frmProgress.Caption = tempCaption
    End If
    
    PackageName = PackageName

    sLog = sLog + Indent + "====== LIP Install: " + PackageName + " ======" + VBA.vbNewLine

    sLog = sLog + Indent + "Looking for package: '" + PackageName + "'" + VBA.vbNewLine
    Set Package = SearchForPackageInStores(PackageName)
    
    ' If no package was found, exit
    If Package Is Nothing Then
        Application.MousePointer = 0
        If Not m_frmProgress Is Nothing Then
            m_frmProgress.Hide
            Set m_frmProgress = Nothing
        End If
        Exit Sub
    End If
    
    
    If Package.Exists("source") Then
        downloadURL = VBA.Replace(Package.Item("source"), "\/", "/") 'Replace \/ with only / since JSON escapes frontslash with a backslash which causes problems with URLs
    Else
        'Handle local source
        If Package.Exists("localsource") Then
            downloadURL = Package.Item("localsource")
            Call InstallFromZip(False, downloadURL)
            Exit Sub
        Else
            downloadURL = BaseURL & PackageStoreApiURL & PackageName & "/download/"  'Use Lundalogik Packagestore if source-node wasn't found
        End If
        
    End If

    If Package.Exists("installPath") Then
        sInstallPath = ThisApplication.TemporaryFolder & Package.Item("installPath") & "\"
    Else
        sInstallPath = ThisApplication.TemporaryFolder & DefaultInstallPath
    End If

    Set Package = Package
    
    
    
    'LJE This is not used anymore
    ''Parse result from store
    'PackageVersion = findNewestVersion(Package.Item("versions"))
    '
    ''Check if package already exsists
    'If Not upgrade Then
    '    If CheckForLocalInstalledPackage(PackageName, PackageVersion) = True Then
    '        Call Lime.MessageBox("Package already installed. If you want to upgrade the package, run command: " & VBA.vbNewLine & VBA.vbNewLine & "Call lip.Install(""" & PackageName & """, True)", vbInformation)
    '        Application.MousePointer = 0
    '        Exit Sub
    '    End If
    'End If
    
    'Install dependecies
    If Package.Exists("dependencies") Then
        If Package("dependencies").Count > 0 Then
            Call updateProgressBar("Installing dependencies", m_progressDouble)
        
            IncreaseIndent
            
            tempProgress = m_progressDouble
            tempCaption = m_frmProgress.Caption
            m_progressDouble = 0
            
            Call InstallDependencies(Package, Simulate)
            
            If m_frmProgress Is Nothing Then
                Set m_frmProgress = New FormProgress
                m_frmProgress.show
            End If
            m_progressDouble = tempProgress
            m_frmProgress.Caption = tempCaption
            
            DecreaseIndent
        End If
    End If
    
    'Download and unzip
    sLog = sLog + Indent + "Downloading '" + PackageName + "' files..." + VBA.vbNewLine
    Dim strDownloadError As String
    strDownloadError = DownloadFile(PackageName, downloadURL, sInstallPath)
    If strDownloadError = "" Then
        Call UnZip(PackageName, sInstallPath)
        sLog = sLog + Indent + "Download complete!" + VBA.vbNewLine
        
               
        If InstallPackageComponents(PackageName, PackageVersion, Package, sInstallPath, Simulate) = False Then
            bOK = False
        End If
    Else
        bOK = False
        sLog = sLog + Indent + strDownloadError
    End If
    
    If bOK Then
        If Simulate Then
            sLog = sLog + Indent + "Simulation of " + PackageName + " done!" + VBA.vbNewLine
        Else
            sLog = sLog + Indent + "Installation of " + PackageName + " done!" + VBA.vbNewLine
        End If
    Else
        sLog = sLog + Indent + "Errors or warnings were raised while installing " + PackageName + ". Please check the log above." + VBA.vbNewLine
    End If

    sLog = sLog + Indent + "===================================" + VBA.vbNewLine
    
    Dim sLogfile As String
    sLogfile = Application.TemporaryFolder & "\" & PackageName & GetCleanTimestamp() & ".txt"
    Open sLogfile For Output As #1
    Print #1, sLog
    Close #1
    
    ' Finish the progress bar and hide it
    If Not m_frmProgress Is Nothing Then
        Call updateProgressBar(VBA.IIf(Simulate, "Simulation", "Installation") & " done!", 99)
        m_frmProgress.Hide
        Set m_frmProgress = Nothing
    End If
    
    If Simulate Then
        Call ThisApplication.Shell(sLogfile)
        If bOK Then
            If VBA.vbYes = Lime.MessageBox("Simulation of installation process completed for package " & PackageName & ". Please check the result in the recently opened logfile." & VBA.vbNewLine & VBA.vbNewLine & "Do you wish to proceed with the installation?", VBA.vbInformation + VBA.vbYesNo + VBA.vbDefaultButton2) Then
                Call lip.Install(PackageName, upgrade, False)
            End If
        Else
            Call Lime.MessageBox("Simulation of installation process completed for package " & PackageName & ". Errors occurred, please check the result in the recently opened logfile and take necessary actions before you try again.")
        End If
    Else
        ' Copy log file to target folder
        Call VBA.FileCopy(sLogfile, LCO.MakeFileName(LCO.MakeFileName(sInstallPath, PackageName), "installationlog_" & GetCleanTimestamp() & ".txt"))
        
        ' Prompt success message (and possibly reminder of manual steps left to do)
        Dim sMsg As String
        sMsg = "Installation process completed for package " & PackageName & "."
        
        If VBA.Dir(LCO.MakeFileName(LCO.MakeFileName(sInstallPath, PackageName), "lisa"), VBA.vbDirectory) <> "" Then
            sMsg = sMsg & VBA.vbCrLf & VBA.vbCrLf & "Please note that there are things that must be manually installed in LISA!"
        End If
        
        If VBA.Dir(LCO.MakeFileName(LCO.MakeFileName(sInstallPath, PackageName), "sql"), VBA.vbDirectory) <> "" Then
            sMsg = sMsg & VBA.vbCrLf & VBA.vbCrLf & "Please note that there is SQL code that must be manually installed!"
        End If
        
        Call Lime.MessageBox(sMsg, VBA.vbInformation + VBA.vbOKOnly)
        
        ' Open folder containing log file and files for manual installation
        Call Application.Shell(sInstallPath + PackageName)
    End If
    
    sLog = ""
    Application.MousePointer = 0
    
    Exit Sub
ErrorHandler:
    If Not m_frmProgress Is Nothing Then
        m_frmProgress.Hide
        Set m_frmProgress = Nothing
    End If
    Call LC_UI.ShowError("lip.Install")
End Sub

'Installs package from a zip-file
Public Sub InstallFromZip(Optional bBrowse As Boolean = True, Optional sZipPath As String = "", Optional Simulate As Boolean = True)
On Error GoTo ErrorHandler
    
    Dim bOK As Boolean
    Dim sInstallPath As String
    Dim tempProgress As Double
    Dim tempCaption As String
    
    bOK = True
    sLog = ""
    Indent = ""
    IndentLenght = "  "
    
    If bBrowse Then
        sZipPath = selectZipFile
        
        ' Just abort if no zip was chosen
        If sZipPath = "" Then
            Exit Sub
        End If
    End If
    
    If Not VerifyLIPInstallation Then
        Call Lime.MessageBox("Verification of LIP failed, installation aborted.", vbCritical)
        Exit Sub
    End If
    
    'Check if valid path
    If sZipPath <> "" Then
        If VBA.Right(sZipPath, 4) = ".zip" Then
            If VBA.Dir(sZipPath) <> "" Then
                Application.MousePointer = 11
                
                ' Initialize the progress bar
                If m_frmProgress Is Nothing Then
                    Set m_frmProgress = New FormProgress
                    m_frmProgress.Caption = VBA.IIf(Simulate, "Simulating", "Installing")
                    m_progressDouble = 0
                End If
                m_frmProgress.show
                
                ' Check if first use of LIP ever
                If VBA.Dir(ThisApplication.WebFolder + "packages.json") = "" Then
                    sLog = sLog + Indent + "No packages.json found, assuming fresh install" + VBA.vbNewLine
                    
                    tempProgress = m_progressDouble
                    tempCaption = m_frmProgress.Caption
                    m_progressDouble = 0
                    
                    Call InstallLIP
                    
                    If m_frmProgress Is Nothing Then
                        Set m_frmProgress = New FormProgress
                        m_frmProgress.show
                    End If
                    m_progressDouble = tempProgress
                    m_frmProgress.Caption = tempCaption
                End If
                
                ' Copy file to actionpads\apps
                Dim PackageName As String
                Dim strArray() As String
                strArray = VBA.Split(sZipPath, "\")
                PackageName = VBA.Split(strArray(UBound(strArray)), ".")(0)
                sLog = sLog + Indent + "====== LIP " & VBA.IIf(Simulate, "Simulation", "Installation") & ": " + PackageName + " ======" + VBA.vbNewLine
                sLog = sLog + Indent + "Copying and unzipping file" + VBA.vbNewLine
                
                If Simulate Then
                    m_frmProgress.Caption = "Simulating " & PackageName
                Else
                    m_frmProgress.Caption = "Installing " & PackageName
                End If
                Call updateProgressBar("Copying and unzipping file", m_progressDouble)
                
                '##TODO: Redo this: the terms app and package have been redefined.
                ' If prefix = app_ or app- then change installpath to /apps else /packages
                If VBA.Left(PackageName, 4) = "app_" Or VBA.Left(PackageName, 4) = "app-" Then
                    sInstallPath = Application.TemporaryFolder & "apps\"
                Else
                    sInstallPath = Application.TemporaryFolder & DefaultInstallPath
                End If
                
                'If apps\packes folder doesn't exist
                If VBA.Dir(sInstallPath, vbDirectory) = "" Then
                    Call VBA.MkDir(sInstallPath)
                End If
                'Copy zip-file to the apps-folder if it's not already there
                If sZipPath <> sInstallPath & PackageName & ".zip" Then
                    Call VBA.FileCopy(sZipPath, sInstallPath & PackageName & ".zip")
                End If
                
                ' Unzip file
                Call UnZip(PackageName, sInstallPath)
    
                ' Get package information from json-file
                Dim Package As Object
                Dim sJson As String
                Dim sLine As String
        
                ' Look for lip.json
                Dim sLipJsonPath As String
                sLipJsonPath = sInstallPath & PackageName & "\" & "lip.json"
                If VBA.Dir(sLipJsonPath) <> "" Then
                    Open sLipJsonPath For Input As #1
                Else
                    Dim sErrorMessage As String
                    sErrorMessage = "Installation failed: could not find any lip.json in the zip-file"
                    sLog = sLog + Indent + sErrorMessage + VBA.vbNewLine
                    Call Application.MessageBox("ERROR: " + sErrorMessage)
                    Application.Shell SaveLogFile(PackageName)
                    If Not m_frmProgress Is Nothing Then
                        m_frmProgress.Hide
                        Set m_frmProgress = Nothing
                    End If
                    Exit Sub
                End If
                
                ' Read JSON from file
                Do Until EOF(1)
                    Line Input #1, sLine
                    sJson = sJson & sLine
                Loop
                Close #1
                Set Package = JSON.parse(sJson)
                
                ' ##TODO: Vad är installPath för inställning?
                If Package.Exists("installPath") Then
                    sInstallPath = ThisApplication.TemporaryFolder & Package.Item("installPath") & "\"
                End If
                
                'Install dependencies
                If Package.Exists("dependencies") Then
                    If Package("dependencies").Count > 0 Then
                        IncreaseIndent
                        
                        tempProgress = m_progressDouble
                        tempCaption = m_frmProgress.Caption
                        m_progressDouble = 0
                        Call updateProgressBar("Installing dependencies", m_progressDouble)
                        
                        Call InstallDependencies(Package, Simulate)
                        
                        If m_frmProgress Is Nothing Then
                            Set m_frmProgress = New FormProgress
                            m_frmProgress.show
                        End If
                        m_progressDouble = tempProgress
                        m_frmProgress.Caption = tempCaption
                        
                        DecreaseIndent
                    End If
                End If
                
                If Not InstallPackageComponents(PackageName, 1, Package, sInstallPath, Simulate) Then
                    bOK = False
                End If
                
                If bOK Then
                    If Simulate Then
                        sLog = sLog + Indent + "Simulation of " + PackageName + " done!" + VBA.vbNewLine
                    Else
                        sLog = sLog + Indent + "Installation of " + PackageName + " done!" + VBA.vbNewLine
                    End If
                Else
                    sLog = sLog + Indent + "Errors or warnings were raised while installing " + PackageName + ". Please check the log above." + VBA.vbNewLine
                End If
    
                sLog = sLog + Indent + "===================================" + VBA.vbNewLine
                
                Dim sLogfile As String
                sLogfile = Application.TemporaryFolder & "\" & PackageName & GetCleanTimestamp() & ".txt"
                Open sLogfile For Output As #1
                Print #1, sLog
                Close #1
                
                ' Finish the progress bar and hide it
                Call updateProgressBar(VBA.IIf(Simulate, "Simulation", "Installation") & " done!", 99)
                m_frmProgress.Hide
                Set m_frmProgress = Nothing
                
                If Simulate Then
                    Call ThisApplication.Shell(sLogfile)
                    If bOK Then
                        If vbYes = Lime.MessageBox("Simulation of installation process completed for package " & PackageName & ". Please check the result in the recently opened logfile." & VBA.vbNewLine & VBA.vbNewLine & "Do you wish to proceed with the installation?", vbInformation + vbYesNo + vbDefaultButton2) Then
                            Call InstallFromZip(False, sZipPath, False)
                        End If
                    Else
                        Call Lime.MessageBox("Simulation of installation process completed for package " & PackageName & ". Errors occurred, please check the result in the recently opened logfile and take necessary actions before you try again.")
                    End If
                Else
                    ' Copy log file to target folder
                    Call VBA.FileCopy(sLogfile, LCO.MakeFileName(LCO.MakeFileName(sInstallPath, PackageName), "installationlog_" & GetCleanTimestamp() & ".txt"))
                    
                    ' Prompt success message (and possibly reminder of manual steps left to do)
                    Dim sMsg As String
                    sMsg = "Installation process completed for package " & PackageName & "."
                    
                    If VBA.Dir(LCO.MakeFileName(LCO.MakeFileName(sInstallPath, PackageName), "lisa"), VBA.vbDirectory) <> "" Then
                        sMsg = sMsg & VBA.vbCrLf & VBA.vbCrLf & "Please note that there are things that must be manually installed in LISA!"
                    End If
                    
                    If VBA.Dir(LCO.MakeFileName(LCO.MakeFileName(sInstallPath, PackageName), "sql"), VBA.vbDirectory) <> "" Then
                        sMsg = sMsg & VBA.vbCrLf & VBA.vbCrLf & "Please note that there is SQL code that must be manually installed!"
                    End If
                    
                    Call Lime.MessageBox(sMsg, VBA.vbInformation + VBA.vbOKOnly)
                    
                    ' Open folder containing log file and files for manual installation
                    Call Application.Shell(sInstallPath + PackageName)
                End If
                
            Else
                Call Lime.MessageBox("Could not find file")
            End If
        Else
            Call Lime.MessageBox("Path must end with .zip")
        End If
    Else
        Call Lime.MessageBox("No path to zip file provided")
    End If
    
    Set m_frmProgress = Nothing
    
    sLog = ""
    Application.MousePointer = 0
    
    Exit Sub
ErrorHandler:
    If Not m_frmProgress Is Nothing Then
        m_frmProgress.Hide
        Set m_frmProgress = Nothing
    End If
    Call LC_UI.ShowError("lip.InstallFromZip")
End Sub

Private Function SaveLogFile(strPackageName As String) As String
    Dim sLogfile As String
    sLogfile = Application.TemporaryFolder & "\" & strPackageName & GetCleanTimestamp() & ".txt"
    Open sLogfile For Output As #1
    Print #1, sLog
    Close #1
    
    SaveLogFile = sLogfile
End Function

'Installs all packages defined in the packages.json file
Public Sub InstallFromPackageFile()
On Error GoTo ErrorHandler
    Dim LocalPackages As Object
    Dim LocalPackageName As Variant

    sLog = sLog + Indent + "Installing dependecies from packages.json file..." + VBA.vbNewLine
    Set LocalPackages = ReadPackageFile().Item("dependencies")
    If LocalPackages Is Nothing Then
        Exit Sub
    End If
    For Each LocalPackageName In LocalPackages.keys
        Call Install(CStr(LocalPackageName), True)
    Next LocalPackageName
Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.InstallFromPackageFile")
End Sub


' ##SUMMARY Returns true if all enforced verifications passed and otherwise false.
Private Function VerifyPackage(PackageName As String, Package As Object) As Boolean
    On Error GoTo ErrorHandler

    ' Verify relations
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    If Package.Item("install").Exists("relations") Then
        sLog = sLog + Indent + "Verifying relations between tables..." + VBA.vbNewLine
        Call updateProgressBar("Verifying relations...", m_progressDouble)
        
        IncreaseIndent
        If Not verifyRelations(Package) Then
            sLog = sLog + Indent + "ERROR: Verification of the relations between tables stated in the package failed!" + VBA.vbNewLine
            DecreaseIndent
            VerifyPackage = False
            Exit Function
        Else
            IncreaseIndent
            sLog = sLog + Indent + "Verification of relations OK" + VBA.vbNewLine
            DecreaseIndent
        End If
        DecreaseIndent
    End If
    
    ' If we end up here, everything went fine.
    VerifyPackage = True

    Exit Function
ErrorHandler:
    Call LC_UI.ShowError("lip.VerifyPackage")
End Function


Private Function InstallPackageComponents(PackageName As String, PackageVersion As Double, Package As Object, InstallPath As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    
    Dim bOK As Boolean
    bOK = True
    
    ' Check if an install node exists in the package
    If Not Package.Exists("install") Then
        sLog = sLog + Indent + "ERROR: The package does not contain an install node." + VBA.vbNewLine
        InstallPackageComponents = False
        Exit Function
    End If
    
    ' Verify content before doing anything else
    If Not VerifyPackage(PackageName, Package) Then
        InstallPackageComponents = False
        Exit Function
    End If
    
    'Install localizations
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    If Package.Item("install").Exists("localize") Then
        sLog = sLog + Indent + "Adding localizations..." + VBA.vbNewLine
        
        Call updateProgressBar("Adding localizations...", m_progressDouble)
        
        IncreaseIndent
        If Not InstallLocalize(Package.Item("install").Item("localize"), Simulate) Then
            bOK = False
        End If
        DecreaseIndent
        
    End If

    'Install VBA
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    If Package.Item("install").Exists("vba") Then
        sLog = sLog + Indent + "Adding VBA modules, forms and classes..." + VBA.vbNewLine
        
        Call updateProgressBar("Adding VBA modules, forms and classes...", m_progressDouble)
        
        IncreaseIndent
        If Not InstallVBAComponents(PackageName, Package.Item("install").Item("vba"), InstallPath, Simulate) Then
            bOK = False
        End If
                
        DecreaseIndent
    End If
    
    
    ' Install tables and fields
    Dim sCreatedTables As String
    Dim sCreatedFields As String
    
    sCreatedTables = ""
    sCreatedFields = ""

    m_progressDouble = m_progressDouble + ProgressBarIncrease
    If Package.Item("install").Exists("tables") Then
        Call updateProgressBar("Adding tables and fields...", m_progressDouble)
        If Not InstallFieldsAndTables(Package.Item("install").Item("tables"), sCreatedTables, sCreatedFields) Then
            bOK = False
        End If
    End If
    
    ' Install relations
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    If Package.Item("install").Exists("relations") Then
        Call updateProgressBar("Adding relations...", m_progressDouble)
        
        If InstallRelations(Package.Item("install").Item("relations"), sCreatedFields) = False Then
            bOK = False
        End If
    End If
    
    ' Copy actionpads
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    If Package.Item("install").Exists("actionpads") Then
        Call updateProgressBar("Copying actionpads...", m_progressDouble)
        If Not InstallActionpads(Package.Item("install").Item("actionpads"), InstallPath & PackageName, Simulate) Then
            bOK = False
        End If
    End If
    
    ' Rollback if only simulation
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    If Simulate Then
        Call updateProgressBar("Rolling back tables and fields...", m_progressDouble)
        
        Call RollbackFieldsAndTables(sCreatedTables, sCreatedFields)
        
    End If

'    If Package.Item("install").Exists("sql") = True Then
'        IncreaseIndent
'        If InstallSQL(Package.Item("install").Item("sql"), PackageName, InstallPath, Simulate) = False Then
'            bOk = False
'        End If
'        DecreaseIndent
'    End If
        
    ' Install files. This is for installation of any kind of files that are not normally part of an add-on.
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    If Package.Item("install").Exists("files") = True Then
        Call IncreaseIndent
        Call updateProgressBar("Installing files...", m_progressDouble)
        If Not InstallFiles(Package.Item("install").Item("files"), PackageName, InstallPath, Simulate) Then
            bOK = False
        End If
        Call DecreaseIndent
    End If
    
    ' Update package file.
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    Call updateProgressBar("Writing to package file...", m_progressDouble)
    
    'Update packages.json
    If WriteToPackagesFile(PackageName, CStr(PackageVersion), Simulate) = False Then
        bOK = False
    End If
    
    m_progressDouble = m_progressDouble + ProgressBarIncrease
    Call updateProgressBar("Ending installation...", m_progressDouble)
    
    If Not EndInstallation Then
        bOK = False
    End If
    
    InstallPackageComponents = bOK
    
Exit Function
ErrorHandler:
    InstallPackageComponents = False
    Call LC_UI.ShowError("lip.InstallPackageComponents")
End Function


' ##SUMMARY Copies the Actionpads defined in the LIP package to the Actionpad folder.
' If a file with the same name already exists, it does not replace that file and prints a warning to the log file.
Private Function InstallActionpads(oJSON As Object, sPackageFolderPath As String, bSimulate As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    sLog = sLog + Indent + "Copying Actionpads..." + VBA.vbNewLine
    
    ' Loop over all actionpad objects in the JSON
    Dim oActionpad As Object
    For Each oActionpad In oJSON
        Call IncreaseIndent
        sLog = sLog + Indent + "Copying Actionpad file """ + oActionpad.Item("fileName") + """..." + VBA.vbNewLine
        
        ' Check if the actionpad already exists
        If VBA.Dir(Application.WebFolder & oActionpad.Item("fileName")) <> "" Then
            Call IncreaseIndent
            sLog = sLog + Indent + "Warning: Actionpad file """ + oActionpad.Item("fileName") + """ already exists and will NOT be replaced!" + VBA.vbNewLine
            Call DecreaseIndent
        Else
            If Not bSimulate Then
                ' Copy the actionpad
                Call VBA.FileCopy(sPackageFolderPath & "\" + oActionpad.Item("relPath"), LCO.MakeFileName(Application.WebFolder, oActionpad.Item("fileName")))
                Call IncreaseIndent
                sLog = sLog + Indent + "Actionpad file """ + oActionpad.Item("fileName") + """ copied to the Actionpads folder. Remember to manually register lbs.html as Actionpad on the affected table." + VBA.vbNewLine
                Call DecreaseIndent
            Else
                ' Just log that it would have been copied.
                Call IncreaseIndent
                sLog = sLog + Indent + "No clash with existing files: Actionpad file """ + oActionpad.Item("fileName") + """ would have been copied to the Actionpads folder." + VBA.vbNewLine
                Call DecreaseIndent
            End If
        End If
        Call DecreaseIndent
    Next oActionpad
    
    InstallActionpads = True

    Exit Function
ErrorHandler:
    InstallActionpads = False
    Call LC_UI.ShowError("lip.InstallActionpads")
End Function


Private Sub InstallDependencies(Package As Object, Simulate As Boolean)
On Error GoTo ErrorHandler
    Dim DependencyName As Variant
    Dim LocalPackage As Object
    sLog = sLog + Indent + "Dependencies found! Installing..." + VBA.vbNewLine
    IncreaseIndent
    For Each DependencyName In Package.Item("dependencies").keys()
        Set LocalPackage = FindPackageLocally(CStr(DependencyName))
        If LocalPackage Is Nothing Then
            sLog = sLog + Indent + "Installing dependency: " + CStr(DependencyName) + VBA.vbNewLine
            Call Install(CStr(DependencyName), Simulate)
        ElseIf CDbl(VBA.Replace(LocalPackage.Item(DependencyName), ".", ",")) < CDbl(VBA.Replace(Package.Item("dependencies").Item(DependencyName), ".", ",")) Then
            Call Install(CStr(DependencyName), True, Simulate)
        Else
        End If
    Next DependencyName
    Call DecreaseIndent
Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.InstallDependencies")
End Sub


Private Function SearchForPackageInStores(PackageName As String) As Object
On Error GoTo ErrorHandler
        
    Set SearchForPackageInStores = SearchForPackageInOnlineStores(PackageName)
    
    If SearchForPackageInStores Is Nothing Then
        Set SearchForPackageInStores = SearchForPackageInLocalStores(PackageName)
        If SearchForPackageInStores Is Nothing Then
            'If we've reached this code, package wasn't found
            Debug.Print Indent + ("Package '" & PackageName & "' not found!")
            Set SearchForPackageInStores = Nothing
        End If
    End If

Exit Function
ErrorHandler:
    Set SearchForPackageInStores = Nothing
    Call LC_UI.ShowError("lip.SearchForPackageInStores")
End Function

'LJE Search for package in online stores
Public Function SearchForPackageInOnlineStores(PackageName As String) As Object
On Error GoTo ErrorHandler
    Dim sJson As String
    Dim oJSON As Object
    Dim oStores As Object
    Dim Path As String
    Dim oStore As Variant
    'LJE changed to onlinestores
    'Set oPackages = ReadPackageFile.Item("stores")
    Set oStores = ReadPackageFile.Item("onlinestores")

    'Loop through packagestores from packages.json
    For Each oStore In oStores


        Path = oStores.Item(oStore)
        sLog = sLog + Indent + ("Looking for package at store '" & oStore & "'") + VBA.vbNewLine
        
        sJson = getJSON(Path + PackageName + "/")

        If sJson <> "" Then
            sJson = VBA.Left(sJson, VBA.Len(sJson) - 1) & ",""source"":""" & oStores.Item(oStore) & """}" 'Add a source node so we know where the package exists
        End If

        Set oJSON = ParseJson(sJson) 'Create a JSON object from the string

        If Not oJSON Is Nothing Then
            If oJSON.Item("error") = "" Then
                'Package found, make sure the install node exists
                If Not oJSON.Item("install") Is Nothing Then
                    sLog = sLog + Indent + ("Package '" & PackageName & "' found on store '" & oStore & "'") + VBA.vbNewLine
                    Set SearchForPackageInOnlineStores = oJSON
                    Exit Function
                Else
                    sLog = sLog + Indent + ("Package '" & PackageName & "' found on store '" & oStore & "' but has no valid install instructions!") + VBA.vbNewLine
                    Set SearchForPackageInOnlineStores = oJSON
                    Exit Function
                End If
            End If
        End If
    Next
    
    'If we've reached this code, package wasn't found
    sLog = sLog + Indent + ("Package '" & PackageName & "' not found!") + VBA.vbNewLine
    Set SearchForPackageInOnlineStores = Nothing

Exit Function
ErrorHandler:
    Set SearchForPackageInOnlineStores = Nothing
    Call LC_UI.ShowError("lip.SearchForPackageInOnlineStores")
End Function


'LJE Search for package in local stores
'Should be a local path where folders are named after packages
'LJE TEST
Public Function SearchForPackageInLocalStores(PackageName As String) As Object
On Error GoTo ErrorHandler
    Dim oStores As Object
    Dim oStore As Variant
    Dim Path As String
    Dim FileSystem As Object
    Dim oJSON As Object
    
    Set oStores = ReadPackageFile.Item("localstores")
    'TODO Test if the oStores is ok
    'TODO Test with multiple local stores
    
    'Loop through localstores from packages.json
    For Each oStore In oStores
        
        Path = oStores.Item(oStore)
        Debug.Print Indent + ("Looking for '" & PackageName & "' at store '" & oStore & "'")
        
        Dim FileSystemObj As Object
        Dim startFolder As Object
        Dim fld As Object
        
        Set FileSystemObj = VBA.CreateObject("Scripting.FileSystemObject")
        'LJE backslash needs to be handled - see trello item.
        'LJE TODO Check if store path is ok
        Set startFolder = FileSystemObj.GetFolder(Path)
        
        
        For Each fld In startFolder.SubFolders
            If LCase(fld.Name) = LCase(PackageName) Then
                Dim sJson As String
                Dim sLine As String
                
                Open fld.Path & "\" & "lip.json" For Input As #1
                        
                Do Until EOF(1)
                    Line Input #1, sLine
                    sJson = sJson & sLine
                Loop
                
                
                If sJson <> "" Then
                    Dim sPathToLocalPackage As String
                    sPathToLocalPackage = VBA.Replace(fld.Path, "\", "\\")
                    sJson = VBA.Left(sJson, VBA.Len(sJson) - 1) & ",""localsource"":""" & sPathToLocalPackage + "\" + fld.Name + """}"   'Add a source node so we know where the package exists
                End If
    
                Close #1
                
                Set oJSON = ParseJson(sJson) 'Create a JSON object from the string
                
                If Not oJSON.Item("install") Is Nothing Then
                    Debug.Print Indent + ("Package '" & PackageName & "' found in local store '" & oStore & "'")
                    Set SearchForPackageInLocalStores = oJSON
                    Exit Function
                Else
                    Debug.Print Indent + ("Package '" & PackageName & "' found in local store '" & oStore & "' but has no valid install instructions!")
                    Set SearchForPackageInLocalStores = Nothing
                    Exit Function
                End If
                
            End If
        Next
    Next
       
    'If we've reached this code, package wasn't found
    Debug.Print Indent + ("Package/App '" & PackageName & "' not found in local stores!")
    Set SearchForPackageInLocalStores = Nothing
    
    Exit Function
ErrorHandler:
    Set SearchForPackageInLocalStores = Nothing
    Call LC_UI.ShowError("lip.SearchForPackageInLocalStores")

End Function

Private Function CheckForLocalInstalledPackage(PackageName As String, PackageVersion As Double) As Boolean
On Error GoTo ErrorHandler
    Dim LocalPackages As Object
    Dim LocalPackage As Object
    Dim LocalPackageVersion As Double
    Dim LocalPackageName As Variant

    Set LocalPackage = FindPackageLocally(PackageName)

    If Not LocalPackage Is Nothing Then
        LocalPackageVersion = CDbl(VBA.Replace(LocalPackage.Item(PackageName), ".", ","))
        If PackageVersion = LocalPackageVersion Then
            sLog = sLog + Indent + "Current version of" + PackageName + " is already installed, please use the upgrade command to reinstall package" + VBA.vbNewLine
            sLog = sLog + Indent + "===================================" + VBA.vbNewLine
            CheckForLocalInstalledPackage = True
            Exit Function
        ElseIf PackageVersion > LocalPackageVersion Then
            sLog = sLog + Indent + "Package " + PackageName + " is already installed, please use the upgrade command to upgrade package from " + Format(LocalPackageVersion, "0.0") + " -> " + Format(PackageVersion, "0.0") + VBA.vbNewLine
            sLog = sLog + Indent + "===================================" + VBA.vbNewLine
            CheckForLocalInstalledPackage = True
            Exit Function
        Else
            sLog = sLog + Indent + "A newer version of " + PackageName + " is already installed. Remote: " + Format(PackageVersion, "0.0") + " ,Local: " + Format(LocalPackageVersion, "0.0") + ". Please use the upgrade command to reinstall package" + VBA.vbNewLine
            sLog = sLog + Indent + "===================================" + VBA.vbNewLine
            CheckForLocalInstalledPackage = True
            Exit Function
        End If
    End If
    CheckForLocalInstalledPackage = False
Exit Function
ErrorHandler:
    Call LC_UI.ShowError("lip.CheckForLocalInstalledPackages")
End Function

Private Function getJSON(sURL As String) As String
On Error GoTo ErrorHandler
    Dim qs As String
    qs = CStr(Rnd() * 1000000#)
    Dim oXHTTP As Object
    Dim s As String
    Set oXHTTP = VBA.CreateObject("MSXML2.XMLHTTP")
    oXHTTP.Open "GET", sURL + "?" + qs, False
    oXHTTP.Send
    getJSON = oXHTTP.responseText
Exit Function
ErrorHandler:
    getJSON = ""
End Function

Private Function ParseJson(sJson As String) As Object
On Error GoTo ErrorHandler
    Dim oJSON As Object
    Set oJSON = JSON.parse(sJson)
    Set ParseJson = oJSON
Exit Function
ErrorHandler:
    Set ParseJson = Nothing
    Call LC_UI.ShowError("lip.parseJSON")
End Function

Private Function findNewestVersion(oVersions As Object) As Double
On Error GoTo ErrorHandler
    Dim NewestVersion As Double
    Dim Version As Variant
    NewestVersion = -1

    For Each Version In oVersions
        If CDbl(VBA.Replace(Version.Item("version"), ".", ",")) > NewestVersion Then
            NewestVersion = CDbl(VBA.Replace(Version.Item("version"), ".", ","))
        End If
    Next Version
    findNewestVersion = NewestVersion
Exit Function
ErrorHandler:
    findNewestVersion = -1
    Call LC_UI.ShowError("lip.findNewestVersion")
End Function

Private Function InstallLocalize(oJSON As Object, Simulate As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bOK As Boolean
    Dim oLocalization As Object
    bOK = True
    
    For Each oLocalization In oJSON
        If Not AddOrCheckLocalize(oLocalization, Simulate) Then
            bOK = False
        End If
    Next oLocalization
    
    ' Reset dictionary to make the new/updated localizations useable.
    Set Localize.dicLookup = Nothing
    
    InstallLocalize = bOK
    
Exit Function
ErrorHandler:
    InstallLocalize = False
    Call LC_UI.ShowError("lip.InstallLocalize")
End Function

Private Function InstallFiles(oJSON As Object, PackageName As String, InstallPath As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim bOK As Boolean
    Dim fso As Object
    Dim FromPath As String
    Dim ToPath As String
    Dim File As Variant
    
    Application.MousePointer = 11
    bOK = True
    

    For Each File In oJSON
        FromPath = InstallPath & PackageName & "\" & File
        ToPath = ThisApplication.WebFolder & File

        If Right(FromPath, 1) = "\" Then
            FromPath = VBA.Left(FromPath, Len(FromPath) - 1)
        End If
        If Right(ToPath, 1) = "\" Then
            ToPath = VBA.Left(ToPath, Len(ToPath) - 1)
        End If
        Set fso = VBA.CreateObject("scripting.filesystemobject")

        fso.CopyFolder Source:=FromPath, Destination:=ToPath
        On Error Resume Next 'It is a beautiful languge
        If Simulate Then
            VBA.Kill ToPath
        Else
            VBA.Kill FromPath
        End If
        On Error GoTo ErrorHandler
    Next File
    
    InstallFiles = bOK

ErrorHandler:
    InstallFiles = False
    sLog = sLog + Indent + ("ERROR: " + Err.Description) + VBA.vbNewLine
    Call LC_UI.ShowError("lip.InstallFiles")
    IncreaseIndent
    DecreaseIndent
End Function

'Private Function InstallSQL(oJSON As Object, PackageName As String, InstallPath As String) As Boolean
'On Error GoTo ErrorHandler
'    Dim bOk As Boolean
'    Dim SQL As Variant
'    Dim Path As String
'    Dim RelPath As String
'
'    bOk = True
'
'    slog=slog+ Indent + "Installing SQL..." +VBA.vbNewLine
'    IncreaseIndent
'    For Each SQL In oJSON
'        RelPath = Replace(SQL.Item("relPath"), "/", "\")
'        Path = InstallPath & PackageName & "\" & RelPath
'        If CreateSQLProcedure(Path, SQL.Item("name"), SQL.Item("type")) = False Then
'            bOk = False
'        End If
'    Next SQL
'    DecreaseIndent
'    InstallSQL = bOk
'Exit Function
'ErrorHandler:
'    InstallSQL = False
'    Call LC_UI.ShowError("lip.InstallSQL")
'End Function
'
'Private Function CreateSQLProcedure(Path As String, Name As String, ProcType As String) As Boolean
'    Dim bOk As Boolean
'    Dim oProc As New LDE.Procedure
'    Dim strSQL As String
'    Dim sLine As String
'    Dim sErrormessage As String
'
'    bOk = True
'    strSQL = ""
'    sErrormessage = ""
'
'    Open Path For Input As #1
'        Do Until EOF(1)
'            Line Input #1, sLine
'            strSQL = strSQL & sLine & VBA.vbNewLine
'        Loop
'        Close #1
'
'        Set oProc = Database.Procedures("csp_lip_installSQL")
'        If Not oProc Is Nothing Then
'            oProc.Parameters("@@sql") = strSQL
'            oProc.Parameters("@@name") = Name
'            oProc.Parameters("@@type") = ProcType
'            oProc.Execute (False)
'
'            sErrormessage = oProc.Parameters("@@errormessage").OutputValue
'
'            If sErrormessage <> "" Then
'                slog=slog+ Indent + (sErrormessage)+VBA.vbNewLine
'                bOk = False
'            Else
'                slog=slog+ Indent + ("'" & Name & "'" & " added.")+VBA.vbNewLine
'            End If
'
'        Else
'            bOk = False
'            Call Lime.MessageBox(GetErrorMessageSQLProcedureNotFound("csp_lip_installSQL"))
'        End If
'
'        CreateSQLProcedure = bOk
'
'Exit Function
'ErrorHandler:
'    CreateSQLProcedure = False
'    Call LC_UI.ShowError("lip.CreateSQLProcedure")
'End Function

Private Function InstallFieldsAndTables(oJSON As Object, ByRef sCreatedTables As String, ByRef sCreatedFields As String) As Boolean
On Error GoTo ErrorHandler
    Dim bOK As Boolean
    Dim table As Object
    Dim oProc As LDE.Procedure
    Dim field As Object
    Dim idtable As Long
    Dim iddescriptiveexpression As Long
    Dim oItem As Variant

    Dim localname_singular As String
    Dim localname_plural As String
    Dim ErrorMessage As String
    Dim warningMessage As String
    
    Dim nbrTables As Integer
    Dim nbrFields As Integer
        
    bOK = True
    
    Application.MousePointer = 11

    sLog = sLog + Indent + "Adding tables and fields..." + VBA.vbNewLine
    
    IncreaseIndent
    
    nbrTables = oJSON.Count
    
    For Each table In oJSON
        localname_singular = ""
        localname_plural = ""
        ErrorMessage = ""
        idtable = -1
        
        Set oProc = Database.Procedures("csp_lip_createtable")
        oProc.Timeout = 299

        If Not oProc Is Nothing Then

            sLog = sLog + Indent + "Add table: " + table.Item("name") + VBA.vbNewLine
            
            Call updateProgressBar("Adding table: " + table.Item("name"), m_progressDouble)
            
            oProc.Parameters("@@tablename").InputValue = table.Item("name")

            'Add localnames singular
            If table.Exists("localname_singular") Then
                For Each oItem In table.Item("localname_singular")
                    If oItem <> "" Then
                        localname_singular = localname_singular + VBA.Trim(oItem) + ":" + VBA.Trim(table.Item("localname_singular").Item(oItem)) + ";"
                    End If
                Next
                oProc.Parameters("@@localname_singular").InputValue = localname_singular
            End If

            'Add localnames plural
            If table.Exists("localname_plural") Then
                For Each oItem In table.Item("localname_plural")
                    If oItem <> "" Then
                        localname_plural = localname_plural + VBA.Trim(oItem) + ":" + VBA.Trim(table.Item("localname_plural").Item(oItem)) + ";"
                    End If
                Next
                oProc.Parameters("@@localname_plural").InputValue = localname_plural
            End If

            Call oProc.Execute(False)

            ErrorMessage = oProc.Parameters("@@errorMessage").OutputValue
            warningMessage = oProc.Parameters("@@warningMessage").OutputValue

            idtable = oProc.Parameters("@@idtable").OutputValue
            iddescriptiveexpression = oProc.Parameters("@@iddescriptiveexpression").OutputValue
            
            If idtable <> -1 Then
                sCreatedTables = sCreatedTables + CStr(idtable) + ";"
            End If

            If warningMessage <> "" Then
                IncreaseIndent
                sLog = sLog + Indent + (warningMessage) + VBA.vbNewLine
                DecreaseIndent
            End If
            
            'If errormessage is set, something went wrong
            If ErrorMessage <> "" Then
                IncreaseIndent
                sLog = sLog + Indent + (ErrorMessage) + VBA.vbNewLine
                bOK = False
                DecreaseIndent
            Else
                sLog = sLog + Indent + ("Table """ & table.Item("name") & """ installed.") + VBA.vbNewLine
            End If

            ' Create fields
            IncreaseIndent
            If table.Exists("fields") Then
                nbrFields = table.Item("fields").Count
                For Each field In table.Item("fields")
                    sLog = sLog + Indent + "Add field: " + table.Item("name") + "." + field.Item("name") + VBA.vbNewLine
                    m_progressDouble = m_progressDouble + (ProgressBarIncrease / nbrTables / nbrFields)
                    Call updateProgressBar("Adding field: " + table.Item("name") + "." + field.Item("name"), m_progressDouble)
                        
                    If AddField(table.Item("name"), field, sCreatedFields) = False Then
                        bOK = False
                    End If
                Next field
            Else
                m_progressDouble = m_progressDouble + (ProgressBarIncrease / nbrTables)
                Call updateProgressBar("Setting table attributes for " + table.Item("name"), m_progressDouble)
            End If

            'Set table attributes(must be done AFTER fields has been created in order to be able to set descriptive expression)
            'Only set attributes if table was created
            If idtable <> -1 Then
                If SetTableAttributes(table, idtable, iddescriptiveexpression) = False Then
                    bOK = False
                End If
            End If

            DecreaseIndent

        Else
            bOK = False
            Call Lime.MessageBox(GetErrorMessageSQLProcedureNotFound("csp_lip_createtable"))
        End If

    Next table
    DecreaseIndent

    Set oProc = Nothing
    
    InstallFieldsAndTables = bOK
    
    Call updateProgressBar("Adding tables and fields done", m_progressDouble)
    
    Exit Function
ErrorHandler:
    Set oProc = Nothing
    InstallFieldsAndTables = False
    sLog = sLog + Indent + ("ERROR: " + Err.Description) + VBA.vbNewLine
    Call LC_UI.ShowError("lip.InstallFieldsAndTables")
    IncreaseIndent
    DecreaseIndent
End Function


Private Function AddField(tableName As String, field As Object, ByRef sCreatedFields As String) As Boolean
On Error GoTo ErrorHandler
    Dim bOK As Boolean
    Dim oProc As New LDE.Procedure
    Dim ErrorMessage As String
    Dim warningMessage As String
    Dim fieldLocalnames As String
    Dim separatorLocalnames As String
    Dim limevalidationtextLocalnames As String
    Dim commentLocalnames As String
    Dim tooltipLocalnames As String
    Dim oItem As Variant
    Dim optionItems As Variant
    Dim idfield As Long
    Dim idcategory As Long
    Dim idstringlocalname As Long
    
    Application.MousePointer = 11
    
    bOK = True
    ErrorMessage = ""
    warningMessage = ""
    fieldLocalnames = ""
    separatorLocalnames = ""
    limevalidationtextLocalnames = ""
    commentLocalnames = ""
    tooltipLocalnames = ""
    idfield = -1
    idcategory = -1
    idstringlocalname = -1
    
    Set oProc = Database.Procedures("csp_lip_createfield")
    oProc.Timeout = 299

    If Not oProc Is Nothing Then
        oProc.Parameters("@@tablename").InputValue = tableName
        oProc.Parameters("@@fieldname").InputValue = field.Item("name")
        oProc.Parameters("@@fieldtype").InputValue = field.Item("attributes").Item("fieldtype")
        oProc.Parameters("@@defaultvalue").InputValue = field.Item("attributes").Item("defaultvalue")
        oProc.Parameters("@@length").InputValue = field.Item("attributes").Item("length")
        oProc.Parameters("@@isnullable").InputValue = field.Item("attributes").Item("isnullable")
        
        Call oProc.Execute(False)
        ErrorMessage = oProc.Parameters("@@errorMessage").OutputValue
        warningMessage = oProc.Parameters("@@warningMessage").OutputValue
        
        idfield = oProc.Parameters("@@idfield").OutputValue
        
        'Log warnings
        If warningMessage <> "" Then
            IncreaseIndent
            sLog = sLog + Indent + (warningMessage) + VBA.vbNewLine
            DecreaseIndent
        End If
        
        'If errormessage is set, something went wrong
        If ErrorMessage <> "" Then
            IncreaseIndent
            sLog = sLog + Indent + (ErrorMessage) + VBA.vbNewLine
            DecreaseIndent
            bOK = False
        End If
        
        If idfield > 0 Then
            sCreatedFields = sCreatedFields + CStr(idfield) + ";"
            
            idcategory = oProc.Parameters("@@idcategory").OutputValue
            idstringlocalname = oProc.Parameters("@@idstringlocalname").OutputValue
            
            sLog = sLog + Indent + ("Field """ & tableName & "." & field.Item("name") & """ installed.") + VBA.vbNewLine
            sLog = sLog + Indent + ("Adding attributes for field: " & tableName & "." & field.Item("name")) + VBA.vbNewLine
            
            ErrorMessage = ""
            warningMessage = ""
            
            Set oProc = Database.Procedures("csp_lip_setfieldattributes")
            oProc.Timeout = 299
            
            If Not oProc Is Nothing Then
                oProc.Parameters("@@idfield").InputValue = idfield
                oProc.Parameters("@@idcategory").InputValue = idcategory
                oProc.Parameters("@@idstringlocalname").InputValue = idstringlocalname
                oProc.Parameters("@@fieldname").InputValue = field.Item("name")
                
                'Add localnames
                If field.Exists("localname") Then
                    For Each oItem In field.Item("localname")
                        If oItem <> "" Then
                            fieldLocalnames = fieldLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("localname").Item(oItem)) + ";"
                        End If
                    Next
                    oProc.Parameters("@@localname").InputValue = fieldLocalnames
                End If
        
                'Add attributes
                If field.Exists("attributes") Then
                    For Each oItem In field.Item("attributes")
                        'Some of the attributes were already set when creating the field
                        If oItem <> "" And oItem <> "defaultvalue" And oItem <> "length" And oItem <> "isnullable" Then
                            If Not oProc.Parameters.Lookup("@@" & oItem, lkLookupProcedureParameterByName) Is Nothing Then
                                oProc.Parameters("@@" & oItem).InputValue = field.Item("attributes").Item(oItem)
                            Else
                                IncreaseIndent
                                sLog = sLog + Indent + ("No support for setting field attribute " & oItem) + VBA.vbNewLine
                                DecreaseIndent
                            End If
                        End If
                    Next
                End If
        
                'Add separator
                If field.Exists("separator") Then
                    For Each oItem In field.Item("separator")
                        separatorLocalnames = separatorLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("separator").Item(oItem)) + ";"
                    Next
                    oProc.Parameters("@@separator").InputValue = separatorLocalnames
                End If
                
                'Add limevalidationtext
                If field.Exists("limevalidationtext") Then
                    For Each oItem In field.Item("limevalidationtext")
                        limevalidationtextLocalnames = limevalidationtextLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("limevalidationtext").Item(oItem)) + ";"
                    Next
                    oProc.Parameters("@@limevalidationtext").InputValue = limevalidationtextLocalnames
                End If
                
                'Add comment
                If field.Exists("comment") Then
                    For Each oItem In field.Item("comment")
                        commentLocalnames = commentLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("comment").Item(oItem)) + ";"
                    Next
                    oProc.Parameters("@@comment").InputValue = commentLocalnames
                End If
                
                'Add tooltip (description)
                If field.Exists("description") Then
                    For Each oItem In field.Item("description")
                        tooltipLocalnames = tooltipLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("description").Item(oItem)) + ";"
                    Next
                    oProc.Parameters("@@description").InputValue = tooltipLocalnames
                End If
        
                Dim strOptions As String
                strOptions = ""
                'Add options
                If field.Exists("options") Then
                    For Each optionItems In field.Item("options")
                        strOptions = strOptions + "["
                        For Each oItem In optionItems
                            strOptions = strOptions + VBA.Trim(oItem) + ":" + VBA.Trim(optionItems.Item(oItem)) + ";"
                        Next
                        strOptions = strOptions + "]"
                    Next
                    oProc.Parameters("@@optionlist").InputValue = strOptions
                End If
                
                Call oProc.Execute(False)
                
                ErrorMessage = oProc.Parameters("@@errorMessage").OutputValue
                warningMessage = oProc.Parameters("@@warningMessage").OutputValue
                
                'Log warnings
                If warningMessage <> "" Then
                    IncreaseIndent
                    sLog = sLog + Indent + (warningMessage) + VBA.vbNewLine
                    DecreaseIndent
                End If
                
                'If errormessage is set, something went wrong
                If ErrorMessage <> "" Then
                    IncreaseIndent
                    sLog = sLog + Indent + (ErrorMessage) + VBA.vbNewLine
                    DecreaseIndent
                    bOK = False
                Else
                    sLog = sLog + Indent + ("Attributes for field """ & tableName & "." & field.Item("name") & """ set.") + VBA.vbNewLine
                End If
            Else
                bOK = False
                Call Lime.MessageBox(GetErrorMessageSQLProcedureNotFound("csp_lip_setfieldattributes"))
            End If
        End If
    Else
        bOK = False
        Call Lime.MessageBox(GetErrorMessageSQLProcedureNotFound("csp_lip_createfield"))
    End If
    Set oProc = Nothing
    AddField = bOK

    Exit Function
ErrorHandler:
    Set oProc = Nothing
    AddField = False
    sLog = sLog + Indent + ("ERROR: " + Err.Description) + VBA.vbNewLine
    Call LC_UI.ShowError("lip.AddField")
    IncreaseIndent
    DecreaseIndent
End Function

Private Function SetTableAttributes(ByRef table As Object, idtable As Long, iddescriptiveexpression As Long) As Boolean
On Error GoTo ErrorHandler

    Dim bOK As Boolean
    Dim oProcAttributes As LDE.Procedure
    Dim oItem As Variant
    Dim ErrorMessage As String
    Dim warningMessage As String
    
    Application.MousePointer = 11
    
    bOK = True
    ErrorMessage = ""
    warningMessage = ""

    If table.Exists("attributes") Then

        Set oProcAttributes = Application.Database.Procedures("csp_lip_settableattributes")
        oProcAttributes.Timeout = 299

        If Not oProcAttributes Is Nothing Then

            sLog = sLog + Indent + "Adding attributes for table: " + table.Item("name") + VBA.vbNewLine

            oProcAttributes.Parameters("@@tablename").InputValue = table.Item("name")
            oProcAttributes.Parameters("@@idtable").InputValue = idtable
            oProcAttributes.Parameters("@@iddescriptiveexpression").InputValue = iddescriptiveexpression

            For Each oItem In table.Item("attributes")
                If oItem <> "" Then
                    If Not oProcAttributes.Parameters.Lookup("@@" & oItem, lkLookupProcedureParameterByName) Is Nothing Then
                        oProcAttributes.Parameters("@@" & oItem).InputValue = table.Item("attributes").Item(oItem)
                    Else
                        sLog = sLog + Indent + ("No support for setting table attribute " & oItem) + VBA.vbNewLine
                    End If
                End If
            Next

            Call oProcAttributes.Execute(False)

            ErrorMessage = oProcAttributes.Parameters("@@errorMessage").OutputValue
            warningMessage = oProcAttributes.Parameters("@@warningMessage").OutputValue
            
            If warningMessage <> "" Then
                sLog = sLog + Indent + (warningMessage) + VBA.vbNewLine
            End If

            'If errormessage is set, something went wrong
            If ErrorMessage <> "" Then
                sLog = sLog + Indent + (ErrorMessage) + VBA.vbNewLine
                bOK = False
            Else
                sLog = sLog + Indent + ("Attributes for table """ & table.Item("name") & """ set.") + VBA.vbNewLine
            End If

        Else
            bOK = False
            Call Lime.MessageBox(GetErrorMessageSQLProcedureNotFound("csp_lip_settableattributes"))
        End If
    End If

    Set oProcAttributes = Nothing
    
    SetTableAttributes = bOK

    Exit Function
ErrorHandler:
    Set oProcAttributes = Nothing
    SetTableAttributes = False
    sLog = sLog + Indent + ("ERROR: " + Err.Description) + VBA.vbNewLine
    Call LC_UI.ShowError("lip.SetTableAttributes")
    IncreaseIndent
    DecreaseIndent
End Function

Private Function DownloadFile(PackageName As String, Path As String, InstallPath As String) As String
On Error GoTo ErrorHandler
    Dim qs As String
    qs = CStr(Rnd() * 1000000#)
    Dim downloadURL As String
    Dim myURL As String
    Dim oStream As Object
    downloadURL = Path + PackageName + "/download/"

    Dim WinHttpReq As Object
    Set WinHttpReq = VBA.CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", downloadURL + "?" + qs, False
    WinHttpReq.Send
    
    DownloadFile = ""
    
    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = VBA.CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile InstallPath + PackageName + ".zip", 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    Else
        DownloadFile = "Couldn't download file from " & downloadURL & vbCrLf & vbCrLf & Err.Description
    End If
    
    Exit Function
ErrorHandler:
    DownloadFile = "Couldn't download file from " & downloadURL & vbCrLf & vbCrLf & Err.Description
End Function

Private Sub UnZip(PackageName As String, InstallPath As String)
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim oApp As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String

    Fname = InstallPath + PackageName + ".zip"
    FileNameFolder = InstallPath & PackageName & "\"

    On Error Resume Next
    Set fso = VBA.CreateObject("scripting.filesystemobject")
    'Delete files
    Call fso.DeleteFile(FileNameFolder & "*.*", True)
    'Delete subfolders
    Call fso.DeleteFolder(FileNameFolder & "*.*", True)

    'Make the normal folder in DefPath
    Call VBA.MkDir(FileNameFolder)

    Set oApp = VBA.CreateObject("Shell.Application")
    oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).Items

    'Delete zip-file
    Call fso.DeleteFile(Fname, True)

    Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.Unzip")
End Sub

Private Function InstallVBAComponents(PackageName As String, VBAModules As Object, InstallPath As String, Simulate As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim bOK As Boolean
    bOK = True
    Dim VBAModule As Variant
    For Each VBAModule In VBAModules
        If Not addModule(PackageName, VBAModule.Item("name"), VBAModule.Item("relPath"), InstallPath, Simulate) Then
            bOK = False
        End If
    Next VBAModule
    
    InstallVBAComponents = bOK
    
    Exit Function
ErrorHandler:
    InstallVBAComponents = False
    Call LC_UI.ShowError("lip.InstallVBAComponents")
End Function

Private Function addModule(PackageName As String, ModuleName As String, RelPath As String, InstallPath As String, Simulate As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim bOK As Boolean
    bOK = True
    Application.MousePointer = 11
    If PackageName <> "" And ModuleName <> "" Then
        Dim VBComps As Object
        Dim Path As String
        Dim tempModuleName As String

        Set VBComps = Application.VBE.ActiveVBProject.VBComponents

        Path = InstallPath + PackageName + "\" + Replace(RelPath, "/", "\")
        
        If VBA.Dir(Path) <> "" Then
            If ComponentExists(ModuleName, VBComps) Then
                If VBA.vbYes = Lime.MessageBox("Do you want to replace existing VBA-module """ & ModuleName & """?", VBA.vbYesNo + VBA.vbDefaultButton2 + VBA.vbQuestion) Then
                    tempModuleName = LCO.GenerateGUID
                    tempModuleName = VBA.Replace(VBA.Mid(tempModuleName, 2, VBA.Len(tempModuleName) - 2), "-", "")
                    tempModuleName = VBA.Left("OLD_" & tempModuleName, 30)
                    
                    If Not Simulate Then
                        VBComps.Item(ModuleName).Name = tempModuleName
                    End If
                    
                    If VBA.vbYes = Lime.MessageBox("Do you want to delete the old module?", VBA.vbYesNo + VBA.vbDefaultButton2 + VBA.vbQuestion) Then
                        If Not Simulate Then
                            Call VBComps.Remove(VBComps.Item(tempModuleName))
                        End If
                    Else
                        Call Lime.MessageBox("Old module is saved with the name """ & tempModuleName & """", VBA.vbInformation)
                        sLog = sLog + Indent + ("Old module is saved with the name """ & tempModuleName & """") + VBA.vbNewLine
                    End If
                    
                    If Not Simulate Then
                        Call Application.VBE.ActiveVBProject.VBComponents.Import(Path)
                    End If
                    sLog = sLog + Indent + "VBA added: " + ModuleName + VBA.vbNewLine
                Else
                    sLog = sLog + Indent + ("Module """ & ModuleName & """ already exists and have not been replaced.") + VBA.vbNewLine
                End If
            Else
                
                If Not Simulate Then
                    Call Application.VBE.ActiveVBProject.VBComponents.Import(Path)
                End If
                sLog = sLog + Indent + "Added " + ModuleName + VBA.vbNewLine
            End If
        Else
            sLog = sLog + Indent + "Module """ & ModuleName & """ can't be added. File does not exists." + VBA.vbNewLine
            bOK = False
            Exit Function
        End If
        
    Else
        sLog = sLog + (Indent + "Detected invalid package or module name while installing """ + RelPath + """") + VBA.vbNewLine
        bOK = False
        Exit Function
    End If
    addModule = bOK
    
    Exit Function
ErrorHandler:
    addModule = False
    sLog = sLog + Indent + ("ERROR: Couldn't add module " + ModuleName + ". " + Err.Description) + VBA.vbNewLine
    Call LC_UI.ShowError("lip.addModule")
End Function

Private Function ComponentExists(ComponentName As String, VBComps As Object) As Boolean
On Error GoTo ErrorHandler
    Dim VBComp As Variant

    For Each VBComp In VBComps
        If VBComp.Name = ComponentName Then
             ComponentExists = True
             Exit Function
        End If
    Next VBComp

    ComponentExists = False

    Exit Function
ErrorHandler:
    Call LC_UI.ShowError("lip.ComponentExists")
End Function

Private Function WriteToPackagesFile(PackageName As String, Version As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim bOK As Boolean
    Dim oJSON As Object
    Dim fs As Object
    Dim a As Object
    Dim Line As Variant
    
    Application.MousePointer = 11
    bOK = True
    Set oJSON = ReadPackageFile

    oJSON.Item("dependencies").Item(PackageName) = Version
    
    If Not Simulate Then
        Set fs = VBA.CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(WebFolder + "packages.json", True)
        For Each Line In Split(PrettyPrintJSON(JSON.toString(oJSON)), vbCrLf)
            Line = VBA.Replace(Line, "\/", "/") 'Replace \/ with only / since JSON escapes frontslash with a backslash which causes problems with packagestores URLs
            a.WriteLine Line
        Next Line
        a.Close
    End If
    
    WriteToPackagesFile = bOK
    Exit Function
ErrorHandler:
    WriteToPackagesFile = False
    sLog = sLog + Indent + ("ERROR: " + Err.Description) + VBA.vbNewLine
    Call LC_UI.ShowError("lip.WriteToPackagesFile")
    IncreaseIndent
    DecreaseIndent
End Function

Private Function PrettyPrintJSON(JSON As String) As String
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim Indent As String
    Dim PrettyJSON As String
    Dim InsideQuotation As Boolean

    For i = 1 To Len(JSON)
        Select Case VBA.Mid(JSON, i, 1)
            Case """"
                PrettyJSON = PrettyJSON + VBA.Mid(JSON, i, 1)
                If InsideQuotation = False Then
                    InsideQuotation = True
                Else
                    InsideQuotation = False
                End If
            Case "{", "["
                If InsideQuotation = False Then
                    Indent = Indent + "    " ' Add to indentation
                    PrettyJSON = PrettyJSON + "{" + VBA.vbCrLf + Indent
                Else
                    PrettyJSON = PrettyJSON + VBA.Mid(JSON, i, 1)
                End If
            Case "}", "["
                If InsideQuotation = False Then
                    Indent = VBA.Left(Indent, Len(Indent) - 4) 'Remove indentation
                    PrettyJSON = PrettyJSON + VBA.vbCrLf + Indent + "}"
                Else
                    PrettyJSON = PrettyJSON + VBA.Mid(JSON, i, 1)
                End If
            Case ","
                If InsideQuotation = False Then
                    PrettyJSON = PrettyJSON + "," + VBA.vbCrLf + Indent
                Else
                    PrettyJSON = PrettyJSON + VBA.Mid(JSON, i, 1)
                End If
            Case Else
                PrettyJSON = PrettyJSON + VBA.Mid(JSON, i, 1)
        End Select
    Next i
    PrettyPrintJSON = PrettyJSON

    Exit Function
ErrorHandler:
    PrettyPrintJSON = ""
    Call LC_UI.ShowError("lip.PrettyPrintJSON")
End Function

Private Function ReadPackageFile() As Object
On Error GoTo ErrorHandler
    Dim sJson As String
    Dim oJSON As Object
    sJson = getJSON(ThisApplication.WebFolder + "packages.json")

    If sJson = "" Then
        sLog = sLog + Indent + "Error: No packages.json found!" + VBA.vbNewLine
        Set ReadPackageFile = Nothing
        Exit Function
    End If

    Set oJSON = JSON.parse(sJson)
    Set ReadPackageFile = oJSON

    Exit Function
ErrorHandler:
    Set ReadPackageFile = Nothing
    Call LC_UI.ShowError("lip.ReadPackageFile")
End Function

Private Function FindPackageLocally(PackageName As String) As Object
On Error GoTo ErrorHandler
    Dim InstalledPackages As Object
    Dim Package As Object
    Dim ReturnDict As New Scripting.Dictionary
    Dim oPackageFile As Object
    Set oPackageFile = ReadPackageFile

    If Not oPackageFile Is Nothing Then

        If oPackageFile.Exists("dependencies") Then
            If oPackageFile("dependencies").Count > 0 Then
                Set InstalledPackages = oPackageFile.Item("dependencies")
                If InstalledPackages.Exists(PackageName) = True Then
                    Call ReturnDict.Add(PackageName, InstalledPackages.Item(PackageName))
                    Set FindPackageLocally = ReturnDict
                    Exit Function
                End If
            End If
        Else
            sLog = sLog + Indent + ("Couldn't find dependencies in packages.json") + VBA.vbNewLine
        End If

    End If

    Set FindPackageLocally = Nothing
    Exit Function
ErrorHandler:
    Set FindPackageLocally = Nothing
    Call LC_UI.ShowError("lip.FindPackageLocally")
End Function


' ##SUMMARY Creates the packages.json file that is needed by LIP.
Public Sub CreateNewPackagesFile()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim tf As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Set tf = fso.CreateTextFile(WebFolder + "packages.json", True)
    tf.WriteLine ("{")
    'LJE VersionHandling
    'TODO write to GitHub
    tf.WriteLine ("    ""lipversion"" : """ & m_sLIPVersion & """,")
    tf.WriteLine ("    ""onlinestores"" : {")
    tf.WriteLine ("        ""PackageStore"" : ""http://api.lime-bootstrap.com/packages/"",")
    tf.WriteLine ("        ""Bootstrap Appstore"" : ""http://api.lime-bootstrap.com/apps/""")
    tf.WriteLine ("    },")
    tf.WriteLine ("    ""localstores"" : {")
    tf.WriteLine ("    },")
    tf.WriteLine ("    ""dependencies"" : {")
    tf.WriteLine ("    }")
    tf.WriteLine ("}")
    
    tf.Close
    
    Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.CreateNewPackagesFile")
End Sub


Public Function GetAllInstalledPackages() As String
On Error GoTo ErrorHandler
    Dim oPackageFile As Object
    Set oPackageFile = ReadPackageFile()

    If Not oPackageFile Is Nothing Then
        GetAllInstalledPackages = JSON.toString(oPackageFile)
    Else
        GetAllInstalledPackages = "{}"
        sLog = sLog + Indent + "Couldn't find dependencies in packages.json" + VBA.vbNewLine
    End If

    Exit Function
ErrorHandler:
    Call LC_UI.ShowError("lip.GetInstalledPackages")
End Function


Public Sub InstallLIP()
On Error GoTo ErrorHandler
    Dim InstallPath As String
    
    ' Initialize log and progress bar
    sLog = ""
    If m_frmProgress Is Nothing Then
        Set m_frmProgress = New FormProgress
        m_frmProgress.show
        m_frmProgress.Caption = "Installing LIP"
        m_progressDouble = 0
    End If
    
    Call updateProgressBar("Creating a new packages.json file", 25)
    sLog = sLog + Indent + "Creating a new packages.json file..." + VBA.vbNewLine
    Call CreateNewPackagesFile
    Dim fso As New FileSystemObject
    InstallPath = ThisApplication.TemporaryFolder & DefaultInstallPath
    If Not fso.FolderExists(InstallPath) Then
        Call fso.CreateFolder(InstallPath)
    End If

    Call updateProgressBar("Installing VBA", 50)
    sLog = sLog + Indent + "Installing JSON-lib..." + VBA.vbNewLine
    Dim strDownloadError
    strDownloadError = DownloadFile("vba_json", BaseURL + AppStoreApiURL, InstallPath)
    If strDownloadError = "" Then
        Call UnZip("vba_json", InstallPath)
                
        Call addModule("vba_json", "JSON", "vba\JSON.bas", InstallPath, False)
        Call addModule("vba_json", "cStringBuilder", "vba\cStringBuilder.cls", InstallPath, False)
    
        Call WriteToPackagesFile("vba_json", "1", False)
    
        sLog = sLog + Indent + "Install of LIP complete!" + VBA.vbNewLine
    Else
        sLog = sLog + Indent + "Could not download the package vba_json from the Appstore: " + BaseURL + AppStoreApiURL
    End If
    Dim sLogfile As String
    sLogfile = Application.TemporaryFolder & "\" & "lip" & GetCleanTimestamp() & ".txt"
    Open sLogfile For Output As #1
    Print #1, sLog
    Close #1
    
    Call updateProgressBar("Installation done!", 99)
    
    m_frmProgress.Hide
    Set m_frmProgress = Nothing
    
    Call Application.Shell(sLogfile)
    
    Application.MousePointer = 0
    
    Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.InstallLIP")
End Sub


' ##SUMMARY Returns true if the localize record specified was either created or updated correctly.
Private Function AddOrCheckLocalize(oLocalization As Object, Simulate As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ' Get keys
    Dim sOwner As String
    Dim sCode As String
    sOwner = oLocalization.Item("owner")
    sCode = oLocalization.Item("code")
    
    ' Build filter and get hit count
    Dim oFilter As New LDE.Filter
    Call oFilter.AddCondition("owner", lkOpEqual, sOwner)
    Call oFilter.AddCondition("code", lkOpEqual, sCode)
    Call oFilter.AddOperator(lkOpAnd)
    
    Dim hitCount As Long
    hitCount = oFilter.hitCount(Database.Classes("localize"))
    
    Dim oItem As Variant
    If hitCount = 0 Then
        ' Create a new record
        If Not Simulate Then
            sLog = sLog + Indent + "Localization " & sOwner & "." & sCode & " not found, creating new!" + VBA.vbNewLine
            
            Dim oRec As New LDE.Record
            Call oRec.Open(Database.Classes("localize"))
'
            For Each oItem In oLocalization
                Call SetRecordPropertyText(oRec, VBA.CStr(oItem), oLocalization(oItem))
            Next oItem
            
            Call oRec.Update
        Else
            sLog = sLog + Indent + "Localization " & sOwner & "." & sCode & " not found, would have created new." + VBA.vbNewLine
        End If
    ElseIf hitCount = 1 Then
        ' Update the existing record found
        If Not Simulate Then
            sLog = sLog + Indent + "Localization " + sOwner + "." + sCode + " was found, updating! " + VBA.vbNewLine
            
            Dim oRecs As New LDE.Records
            Call oRecs.Open(Database.Classes("localize"), oFilter)
            
            For Each oItem In oLocalization
                Call SetRecordPropertyText(oRecs(1), VBA.CStr(oItem), oLocalization(oItem))
            Next oItem
            
            Call oRecs.Update
        Else
            sLog = sLog + Indent + "Localization " + sOwner + "." + sCode + " was found, would have updated. " + VBA.vbNewLine
        End If
    Else
        ' Error, multiple hits on key owner.code.
        sLog = sLog + Indent + "ERROR: There are multiple copies of " & sOwner & "." & sCode & ". Fix this and try again."
        AddOrCheckLocalize = False
        Exit Function
    End If
    
    AddOrCheckLocalize = True
    
    Exit Function
ErrorHandler:
    sLog = sLog + Indent + "ERROR: An error occurred while validating or adding localizations: " + Err.Description + VBA.vbNewLine
    AddOrCheckLocalize = False
End Function


' ##SUMMARY Tries to add a text value to the specified property on a record.
' If there is no field with the specified name it just continues without reporting any errors.
Private Sub SetRecordPropertyText(oRec As LDE.Record, sPropertyName As String, sText As String)
    On Error GoTo ErrorHandler
    
    If oRec.Fields.Exists(sPropertyName) Then
        oRec.Value(sPropertyName) = sText
    End If

    Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.SetRecordPropertyText")
End Sub


Private Sub IncreaseIndent()
On Error GoTo ErrorHandler
    Indent = Indent + IndentLenght
    Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.IncreaseIndent")
End Sub

Private Sub DecreaseIndent()
On Error GoTo ErrorHandler

    If VBA.Len(Indent) - VBA.Len(IndentLenght) > 0 Then
        Indent = VBA.Left(Indent, VBA.Len(Indent) - VBA.Len(IndentLenght))
    Else
        Indent = ""
    End If
    
    Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.DecreaseIndent")
End Sub

Private Function InstallRelations(oJSON As Object, sCreatedFields As String) As Boolean
On Error GoTo ErrorHandler
    Dim bOK As Boolean
    Dim relation As Object
    Dim oProc As LDE.Procedure

    Dim ErrorMessage As String
    Dim warningMessage As String
    
    Dim nbrRelations As Integer
        
    
    bOK = True
    Application.MousePointer = 11
    sLog = sLog + Indent + "Adding relations..." + VBA.vbNewLine
    IncreaseIndent
    
    For Each relation In oJSON
        nbrRelations = oJSON.Count
        ErrorMessage = ""
        warningMessage = ""

        Set oProc = Database.Procedures("csp_lip_addrelations")
        
        If Not oProc Is Nothing Then
            oProc.Timeout = 299
            
            sLog = sLog + Indent + "Add relation between: " + relation.Item("table1") + "." + relation.Item("field1") + " and " + relation.Item("table2") + "." + relation.Item("field2") + VBA.vbNewLine
            m_progressDouble = m_progressDouble + (ProgressBarIncrease / nbrRelations)
            Call updateProgressBar("Add relation between: " + relation.Item("table1") + "." + relation.Item("field1") + " and " + relation.Item("table2") + "." + relation.Item("field2"), m_progressDouble)
            
            oProc.Parameters("@@table1").InputValue = relation.Item("table1")
            oProc.Parameters("@@field1").InputValue = relation.Item("field1")
            oProc.Parameters("@@table2").InputValue = relation.Item("table2")
            oProc.Parameters("@@field2").InputValue = relation.Item("field2")
            oProc.Parameters("@@createdfields").InputValue = sCreatedFields

            Call oProc.Execute(False)
            ErrorMessage = oProc.Parameters("@@errorMessage").OutputValue
            warningMessage = oProc.Parameters("@@warningMessage").OutputValue
            
            IncreaseIndent
            If warningMessage <> "" Then
                sLog = sLog + Indent + warningMessage + VBA.vbNewLine
            End If
            
            'If errormessage is set, something went wrong
            If ErrorMessage <> "" Then
                sLog = sLog + Indent + ErrorMessage + VBA.vbNewLine
                bOK = False
            End If
            
            If ErrorMessage = "" And warningMessage = "" Then
                sLog = sLog + Indent + "Relation between: " + relation.Item("table1") + "." + relation.Item("field1") + " and " + relation.Item("table2") + "." + relation.Item("field2") + " created." + VBA.vbNewLine
            End If
            DecreaseIndent
        Else
            bOK = False
            Call Lime.MessageBox(GetErrorMessageSQLProcedureNotFound("csp_lip_addrelations"))
        End If
    Next relation
    
    DecreaseIndent
    Set oProc = Nothing
    InstallRelations = bOK

    Exit Function
ErrorHandler:
    Set oProc = Nothing
    InstallRelations = False
    sLog = sLog + Indent + ("ERROR: " + Err.Description) + VBA.vbNewLine
    Call LC_UI.ShowError("lip.InstallRelations")
End Function

Private Function RollbackFieldsAndTables(sCreatedTables As String, sCreatedFields As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim oProc As New LDE.Procedure
    Set oProc = Database.Procedures("csp_lip_removetablesandfields")
    oProc.Timeout = 299
    
    If Not oProc Is Nothing Then
        If sCreatedFields <> "" Then
            Dim fieldArray() As String
            fieldArray() = VBA.Split(sCreatedFields, ";")
            
            For i = UBound(fieldArray) - 1 To LBound(fieldArray) Step -1
                oProc.Parameters("@@idfield") = VBA.CLng(fieldArray(i))
                Call oProc.Execute(False)
            Next i
        End If
        
        If sCreatedTables <> "" Then
            Dim tableArray() As String
            tableArray() = VBA.Split(sCreatedTables, ";")
            For i = UBound(tableArray) - 1 To LBound(tableArray) Step -1
                oProc.Parameters("@@idtable") = VBA.CLng(tableArray(i))
                Call oProc.Execute(False)
            Next i
        End If
    Else
        Call Lime.MessageBox(GetErrorMessageSQLProcedureNotFound("csp_lip_removetablesandfields"))
        RollbackFieldsAndTables = False
        Exit Function
    End If
    
    RollbackFieldsAndTables = True
Exit Function
ErrorHandler:
    Call LC_UI.ShowError("lip.RollbackFieldsAndTables")
End Function

'LJE 20160212 Check if a new version of LIP exists
Public Sub UpdateLIPOnNewVersion()
On Error GoTo ErrorHandler
    Dim Package As Object
    Dim PackageVersion As Double
    Dim downloadURL As String
    Dim InstallPath As String
    Dim PackageName As String
    
    Dim oPackageFile As Object
    Set oPackageFile = ReadPackageFile
    
    Indent = ""
    IndentLenght = "  "
    
    PackageName = "lip"
    Debug.Print Indent + "Checking version for LIP"
    Set Package = SearchForPackageInStores("lip")
    
    If Package Is Nothing Then
        Exit Sub
    End If
    
    PackageVersion = findNewestVersion(Package.Item("versions"))
    If PackageVersion > CDbl(VBA.Replace(oPackageFile.Item("lipversion"), ".", ",")) Then
        Debug.Print Indent + "Newer version of lip found"
        
        Dim VBComps As Object
        Dim Path As String
        Dim tempModuleName As String
        
        Set VBComps = Application.VBE.ActiveVBProject.VBComponents
        'LJE TEST
        'VBComps.Item("lip").Name = "lip_old"
        'Call Application.VBE.ActiveVBProject.VBComponents.Import("C:\Temp\LocalStore\lip\Install\VBA\lip.bas")
        
        'LJE TODO Update packages.json with new version
        oPackageFile.Item("lipversion") = VBA.Replace(PackageVersion, ",", ".")
        'LJE TEST
        'Call lip.RemoveModule("lip_old")
        Debug.Print Indent + "LIP updated"
    End If
    Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.UpdateLIPOnNewVersion")
End Sub
'LJE 20160212 Upgrade LIP if new version exists
Private Sub UpdateLIP()
On Error GoTo ErrorHandler
'Q: How to handle the remove of lip.bas.
'Separate lip functions in separate modules, an interface with functions which calls another bas which can be replaced.

'1. Replace lip.bas
'2. Replace csp (this is done manually now)
'3. Tell user what happened and what needs to be done.

 Dim VBComps As Object
 Dim Path As String
 Dim tempModuleName As String

 Set VBComps = Application.VBE.ActiveVBProject.VBComponents
 VBComps.Item("lip").Name = "lip_old"

 Call Application.VBE.ActiveVBProject.VBComponents.Import("C:\Temp\LocalStore\lip\Install\VBA\lip.bas")
 
 'LJE TODO Update packages.json with new version

 Call lip.RemoveModule("lip_old")

'Call VBComps.Remove(VBComps.Item(tempModuleName)
 Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.UpdateLIP")
End Sub

'LJE Remove temporary lip.bas after update
Private Sub RemoveModule(sModuleName As String)
Dim VBComps As Object
On Error GoTo ErrorHandler

Set VBComps = Application.VBE.ActiveVBProject.VBComponents

Call VBComps.Remove(VBComps.Item(sModuleName))
Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.RemoveModule")
End Sub

'LJE TODO Refactor with helper method to write json
Public Sub SetLipVersionInPackageFile(sVersion As String)
On Error GoTo ErrorHandler
'    Open ThisApplication.WebFolder & DefaultInstallPath & PackageName & "\" & "packages.json" For Input As #1
'
'            ElseIf VBA.Dir(ThisApplication.WebFolder & DefaultInstallPath & PackageName & "\" & "lip.json") <> "" Then
'                Open ThisApplication.WebFolder & DefaultInstallPath & PackageName & "\" & "lip.json" For Input As #1
'
'            Else
'                Debug.Print (Indent + "Installation failed: could not find any lip.json in the zip-file")
'                Exit Sub
'            End If
'
'            Do Until EOF(1)
'                Line Input #1, sLine
'                sJSON = sJSON & sLine
'            Loop
'
'            Close #1
    Exit Sub
ErrorHandler:
    Call LC_UI.ShowError("lip.SetLipVersionInPackageFile")
End Sub

Private Function EndInstallation() As Boolean
On Error GoTo ErrorHandler
    Dim bOK As Boolean
    bOK = True
        
    Dim oProc As LDE.Procedure

    Set oProc = Database.Procedures("csp_lip_endinstallation")
    oProc.Timeout = 299

    If Not oProc Is Nothing Then
        Call oProc.Execute(False)
    Else
        bOK = False
        Call Lime.MessageBox(GetErrorMessageSQLProcedureNotFound("csp_lip_endinstallation"))
    End If

    Set oProc = Nothing
    EndInstallation = bOK
    
Exit Function
ErrorHandler:
    Set oProc = Nothing
    EndInstallation = False
    Call LC_UI.ShowError("lip.EndInstallation")
End Function

Private Sub updateProgressBar(sMessage As String, dblProgress As Double)
    On Error GoTo ErrorHandler
    
    If Not m_frmProgress Is Nothing Then
        m_frmProgress.Title.Caption = sMessage
        m_frmProgress.Progress.Width = m_frmProgress.Width / (100 / (dblProgress + 1))
    End If
    
    Exit Sub
ErrorHandler:
    If Not m_frmProgress Is Nothing Then
        m_frmProgress.Hide
        Set m_frmProgress = Nothing
    End If
    Call LC_UI.ShowError("lip.updateProgressBar")
End Sub

'Helper function to get LIP version from packages.json.
Public Function GetInstalledLIPVersion() As String
    Dim bOK As Boolean
    Dim oJSON As Object
    Dim fs As Object
    Dim a As Object
    Dim Line As Variant
    
    On Error GoTo ErrorHandler
    
    Set oJSON = ReadPackageFile
    
    GetInstalledLIPVersion = oJSON.Item("lipversion")
    
    Set oJSON = Nothing
    Debug.Print GetInstalledLIPVersion
Exit Function

ErrorHandler:
    
    Call LC_UI.ShowError("lip.GetInstalledLIPVersion")

End Function


' ##SUMMARY Shows a file dialog where the user can select a zip file.
Private Function selectZipFile() As String
    On Error GoTo ErrorHandler

    Dim fileDialog As New LCO.FileOpenDialog
    fileDialog.Filter = "Zip-file (*.zip) | *.zip"
    fileDialog.AllowMultiSelect = False
    Call fileDialog.show
    
    selectZipFile = fileDialog.FileName

    Exit Function
ErrorHandler:
    Call LC_UI.ShowError("lip.selectZipFile")
End Function



' ##SUMMARY Verify that relations in lip.json will not corrupt the database.
' Either the fields on both sides of a relation do not exist or they both exist
' and in that case they must be linked to eachother as stated in the LIP package.
    
Private Function verifyRelations(Package As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim oRelation As Object
    Dim bField1exists As Boolean
    Dim bField2exists As Boolean
    If Package.Item("install").Exists("relations") Then
        ' Package has relations, loop over them.
        
        For Each oRelation In Package.Item("install").Item("relations")
            sLog = sLog + Indent + "Verifying relation between " + oRelation.Item("table1") + "." + oRelation.Item("field1") + " and " _
                    + oRelation.Item("table2") + "." + oRelation.Item("field2") + "..." + VBA.vbNewLine
            bField1exists = fieldExists(oRelation.Item("table1"), oRelation.Item("field1"))
            bField2exists = fieldExists(oRelation.Item("table2"), oRelation.Item("field2"))
            If bField1exists And bField2exists Then
                ' Both fields exist: Make sure they are related to eachother
                If Not fieldsAreRelated(Database.Classes(oRelation.Item("table1")).Fields(oRelation.Item("field1")), _
                                        Database.Classes(oRelation.Item("table2")).Fields(oRelation.Item("field2"))) Then
                    IncreaseIndent
                    sLog = sLog + Indent + "ERROR: Both fields already exist but they are not related to eachother!" + VBA.vbNewLine
                    DecreaseIndent
                    verifyRelations = False
                    Exit Function
                End If
            ElseIf (bField1exists And Not bField2exists) Or (Not bField1exists And bField2exists) Then
                ' Only one of the fields exists which is not OK.
                IncreaseIndent
                sLog = sLog + Indent + "ERROR: One of the fields already exists!" + VBA.vbNewLine
                DecreaseIndent
                verifyRelations = False
                Exit Function
            End If
        Next oRelation
    End If

    ' If we arrive here, all is well
    verifyRelations = True

    Exit Function
ErrorHandler:
    verifyRelations = False
    Call LC_UI.ShowError("lip.verifyRelations")
End Function


' ##SUMMARY Returns true if the specified field exists in the database and otherwise false.
Private Function fieldExists(tableName As String, fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    If Database.Classes.Exists(tableName) Then
        fieldExists = Database.Classes(tableName).Fields.Exists(fieldName)
    Else
        fieldExists = False
    End If
    
    Exit Function
ErrorHandler:
    fieldExists = False
    Call LC_UI.ShowError("lip.fieldExists")
End Function


' ##SUMMARY Returns true if the two fields specified are relation fields/tabs and are related to eachother.
Private Function fieldsAreRelated(f1 As LDE.field, f2 As LDE.field) As Boolean
    On Error GoTo ErrorHandler

    ' Check if relation fields
    If (Not isRelationField(f1)) Or (Not isRelationField(f2)) Then
        fieldsAreRelated = False
        Exit Function
    End If
    
    ' Check if related to eachother
    If (Not f1.LinkedField Is f2) Or (Not f2.LinkedField Is f1) Then
        fieldsAreRelated = False
        Exit Function
    End If
    
    fieldsAreRelated = True

    Exit Function
ErrorHandler:
    fieldsAreRelated = False
    Call LC_UI.ShowError("lip.fieldsAreRelated")
End Function


' ##SUMMARY Returns true if the specified field is either a relation field or tab.
Private Function isRelationField(f As LDE.field) As Boolean
    On Error GoTo ErrorHandler
    
    isRelationField = ((f.Type And lkFieldTypeLink) = lkFieldTypeLink) Or ((f.Type And lkFieldTypeMultiLink) = lkFieldTypeMultiLink)

    Exit Function
ErrorHandler:
    isRelationField = False
    Call LC_UI.ShowError("lip.isRelationField")
End Function


' ##SUMMARY Returns a timestamp where all characters except digits have been removed.
Private Function GetCleanTimestamp() As String
    On Error GoTo ErrorHandler
    
    Dim sResult As String
    sResult = VBA.Now
    
    ' Replace all special characters that are not approved in file names in Windows
    sResult = VBA.Replace(sResult, "\", "")
    sResult = VBA.Replace(sResult, "/", "")
    sResult = VBA.Replace(sResult, ":", "")
    sResult = VBA.Replace(sResult, "*", "")
    sResult = VBA.Replace(sResult, "?", "")
    sResult = VBA.Replace(sResult, """", "")
    sResult = VBA.Replace(sResult, "<", "")
    sResult = VBA.Replace(sResult, ">", "")
    sResult = VBA.Replace(sResult, "|", "")
    
    ' Replace all additional unwanted characters that can be part of a timestamp in different locales
    sResult = VBA.Replace(sResult, "-", "")
    sResult = VBA.Replace(sResult, " ", "")
    sResult = VBA.Replace(sResult, ".", "")
    
    GetCleanTimestamp = sResult

    Exit Function
ErrorHandler:
    Call LC_UI.ShowError("lip.GetCleanTimestamp")
End Function


' ##SUMMARY Returns the error message to prompt to a user when a SQL procedure was not found.
Private Function GetErrorMessageSQLProcedureNotFound(sProcedureName As String) As String
    On Error GoTo ErrorHandler

    GetErrorMessageSQLProcedureNotFound = Application.FormatString("Could not find SQL procedure '%1'." & VBA.vbCrLf & VBA.vbCrLf & "Please make sure the procedure exists in the database, run lsp_setdatabasetimestamp and lsp_refreshldc and restart the LDC.", sProcedureName)

    Exit Function
ErrorHandler:
    GetErrorMessageSQLProcedureNotFound = ""
    Call LC_UI.ShowError("lip.GetErrorMessageSQLProcedureNotFound")
End Function

Private Function VerifyLIPInstallation() As Boolean
On Error GoTo ErrorHandler

    Dim sVerifySQL As String
    
    VerifyLIPInstallation = True
    
    sVerifySQL = VerifySQL
    
    If sVerifySQL <> "" Then
        Call Lime.MessageBox("The following SQL procedures were not found:" & vbNewLine & vbNewLine & sVerifySQL & vbNewLine & "Please make sure they are installed and accessible.", vbCritical)
        VerifyLIPInstallation = False
    End If
    
Exit Function
ErrorHandler:
    Call LC_UI.ShowError("lip.VerifyLIPInstallation")
End Function

Private Function VerifySQL() As String
On Error GoTo ErrorHandler

    Dim sProcName As Variant
    Dim sProcNames As Variant
    Dim sNotFound As String
    Dim oProc As LDE.Procedure
    Dim bExists As Boolean
    
    sProcNames = Array("csp_lip_addrelations", _
        "csp_lip_settableattributes", _
        "csp_lip_setfieldattributes", _
        "csp_lip_removetablesandfields", _
        "csp_lip_endinstallation", _
        "csp_lip_createtable", _
        "csp_lip_createfield")
    
    For Each sProcName In sProcNames
    
        bExists = False
        
        For Each oProc In Application.Database.Procedures
            If oProc.Name = sProcName Then
                bExists = True
                Exit For
            End If
        Next oProc
        
        If Not bExists Then
            sNotFound = sNotFound + sProcName + vbNewLine
        End If
        
    Next sProcName
    
    VerifySQL = sNotFound

Exit Function

ErrorHandler:
    Call LC_UI.ShowError("lip.VerifySQL")
End Function
