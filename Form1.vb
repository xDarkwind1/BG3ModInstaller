Option Explicit On
Imports Scripting

Public Class frmMain

    'Loads settings file on open
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Get appdata folder if it exists
        Dim objFSO As New FileSystemObject, sPath As String
        sPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\DBG3MI\"
        If Not objFSO.FolderExists(sPath) Then
            objFSO.CreateFolder(sPath)
        End If

        'read the FolderPaths text file and push to textboxes
        Call readPathsFile(sPath)
    End Sub

    'On button click, sets the modsFolder to the chosen path
    Private Sub cmdSetModPath_Click(sender As Object, e As EventArgs) Handles cmdSetModPath.Click
        Dim sPath As String
        sPath = GetFolderPath("Mods")
        If sPath <> vbNullString Then
            txtModsFolder.Text = sPath
        End If
    End Sub

    'On button click, sets the game folder to the chosen path
    Private Sub cmdSetGamePath_Click(sender As Object, e As EventArgs) Handles cmdSetGamePath.Click
        Dim sPath As String
        sPath = GetFolderPath("Game")
        If sPath <> vbNullString Then
            txtModsFolder.Text = sPath
        End If
    End Sub

    'On button click, sets the Backup  folder to the chosen path
    Private Sub cmdSetBackupPath_Click(sender As Object, e As EventArgs) Handles cmdSetBackupPath.Click
        Dim sPath As String
        sPath = GetFolderPath("Backup")
        If sPath <> vbNullString Then
            txtBackupFolder.Text = sPath
        End If
    End Sub

    'On button click, begins installing files from Mods folder to Game folder, backing up overwritten initial game files to Backup folder
    Private Sub cmdInstall_Click(sender As Object, e As EventArgs) Handles cmdInstall.Click

        'Confirm directories are OK and get them
        Dim sPaths As List(Of String), sModsFolder As String, sBackupFolder As String, sGameFolder As String
        sPaths = TestPaths()
        If sPaths.Count <> 3 Then Exit Sub
        sGameFolder = sPaths.Item(0)
        sModsFolder = sPaths.Item(1)
        sBackupFolder = sPaths.Item(2)

        'confirm installation
        If Not MsgBox("Copy mods from " & sModsFolder & " into " & sGameFolder & "?", vbYesNo + vbQuestion, "Confirm Install") = vbYes Then Exit Sub

        'Get list of files to install
        Dim sFiles() As String
        sFiles = getInstallationFiles(sModsFolder)

        'Exit if no files to install and warn the user
        If sFiles.Length = 0 Then
            MsgBox("No files found in the supplied directory. Is the mods folder correct?", vbOK + vbExclamation, "No File Found")
            Exit Sub
        End If

        'Saves list of installation files
        If Not saveInstallationFiles(sBackupFolder, sFiles) Then Exit Sub

        'Save copy of files that will be overwritten
        If Not saveOriginalFiles(sGameFolder, sBackupFolder, sFiles) Then Exit Sub

        'Copy over mod files to directory
        If Not copyModFiles(sModsFolder, sGameFolder) Then Exit Sub

        'Save the used paths to appdata
        If Not savePaths(sGameFolder, sModsFolder, sBackupFolder) Then Exit Sub

        'Inform of success and provide option to exit
        If MsgBox("Installation completed successfully! Close the installer?", vbYesNo + vbQuestion, "Close Installer?") Then
            Application.Exit()
        End If

    End Sub

    'On button click, removes installed files based on saved files from Game folder. Installs files from Backup folder to Game folder.
    Private Sub cmdUninstall_Click(sender As Object, e As EventArgs) Handles cmdUninstall.Click

        Dim sPaths As List(Of String), sModsFolder As String, sBackupFolder As String, sGameFolder As String
        sPaths = TestPaths()
        If sPaths.Count <> 3 Then Exit Sub
        sGameFolder = sPaths.Item(0)
        sModsFolder = sPaths.Item(1)
        sBackupFolder = sPaths.Item(2)

        'Confirm uninstalltion
        If Not MsgBox("Remove mods files (defined in " + sBackupFolder + ") from " + sGameFolder + "?", vbYesNo + vbQuestion, "Confirm Uninstall") = vbYes Then Exit Sub

        'Get list of files to uninstall
        Dim sFiles As New List(Of String)
        sFiles = readInstallationFiles(sBackupFolder)

        'Remove mod files from game folder
        'If Not deleteModFiles(sGameFolder, sFiles) Then Exit Sub

        'Restore original files
        'If Not restoreGameFiles(sGameFolder, sBackupFolder, sFiles) Then Exit Sub

        'Inform of success and provide option to exit
        If MsgBox("Uninstalled successfully. Close the installer?", vbYesNo + vbQuestion, "Close Installer?") Then
            Application.Exit()
        End If

    End Sub

    'Closes the installer
    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Application.Exit()
    End Sub

    'Populates textboxes with data from saved Paths file.
    Private Sub readPathsFile(ByVal sPath As String)
        Dim objFSO As New FileSystemObject
        If Not objFSO.FolderExists(sPath) Then
            Err.Raise(vbObjectError + 514, , "Invalid file path argument when reading the Paths file.")
            Exit Sub
        End If

        If objFSO.FileExists(sPath & "FolderPaths.txt") Then
            Dim objReader As New System.IO.StreamReader(sPath & "FolderPaths.txt")
            Dim sTextLine As String, lIndex As Long
            For lIndex = 1 To 3
                If objReader.Peek() <> -1 Then Exit For
                sTextLine = objReader.ReadLine()
                Select Case lIndex
                    Case 1
                        txtModsFolder.Text = sTextLine
                    Case 2
                        txtGameFolder.Text = sTextLine
                    Case 3
                        txtBackupFolder.Text = sTextLine
                End Select
            Next lIndex
        End If
    End Sub

    'Returns a folder path based on folderdialog
    Public Shared Function GetFolderPath(ByVal sPathType As String) As String
        Dim sPath As String, objFB As New myFolderBrowser With {
            .Description = "Select the " & sPathType & " folder...",
            .StartLocation = myFolderBrowser.enuFolderBrowserFolder.Desktop
        }
        If objFB.ShowBrowser = DialogResult.OK Then
            sPath = objFB.Path
        Else
            sPath = String.Empty
        End If
        Return sPath
    End Function

    'Returns an array of file path names relative to sModsFolder containing each file in the subdirectories
    Private Function getInstallationFiles(ByVal sModsFolder As String) As String()

        'Check nonexistant folder
        Dim objFSO As New FileSystemObject
        If Not objFSO.FolderExists(sModsFolder) Then
            Return Split("")
            Exit Function
        End If

        'recurse files in folder
        Dim objFolder As Folder, lPathLength As Long
        objFolder = objFSO.GetFolder(sModsFolder)
        Return RecurseFolderFiles(objFolder, lPathLength)
    End Function

    'Saves a text list of each file installed
    Private Function saveInstallationFiles(ByVal sBackupFolder As String, ByVal sFiles() As String) As Boolean
        Dim objFSO As New FileSystemObject, sFilesList As New List(Of String)
        sFilesList = sFiles.ToList

        Using objSW As New System.IO.StreamWriter(sBackupFolder & "\InstalledFiles.txt")
            For Each sFile As String In sFilesList
                objSW.WriteLine(sFile)
            Next
            objSW.Flush()
            objSW.Close()
        End Using
        Return True
    End Function

    'Reads the 
    Private Function readInstallationFiles(ByVal sBackupFolder As String) As List(Of String)
        Dim objFSO As New FileSystemObject, sFilesList As New List(Of String)

        If Not objFSO.FileExists(sBackupFolder & "\InstalledFiles.txt") Then
            Return sFilesList
        End If
        Using objSR As New System.IO.StreamReader(sBackupFolder & "\InstalledFiles.txt")
            While objSR.EndOfStream = False
                sFilesList.Add(objSR.ReadLine)
            End While
        End Using
        Return sFilesList
    End Function

    Private Function saveOriginalFiles(ByVal sGameFolder As String, ByVal sBackupFolder As String, ByVal sFiles() As String) As Boolean

        'Push files to list. Loop thru all
        Dim sFileList As New List(Of String), sFile As String
        sFileList = sFiles.ToList
        For Each sFile In sFileList

            'Split file name to get folders in path. Remove the file name from the list.
            Dim sFolders As New List(Of String)
            sFolders = sFile.Split("\").ToList
            sFolders.RemoveAt(sFolders.Count - 1)

            'Check for each folder in backup directory. Create folder if needed
            Dim objFSO As New FileSystemObject, sFolder As String, sDir As String = sBackupFolder & "\"
            For Each sFolder In sFolders
                sDir += sFolder + "\"
                If Not objFSO.FolderExists(sDir) Then
                    objFSO.CreateFolder(sDir)
                End If
            Next

            'Copy file to proper directory
            objFSO.CopyFile(sGameFolder + sFile, sDir, True)
        Next
        Return True
    End Function

    'Copies mod files into game folder
    Private Function copyModFiles(ByVal sModsFolder As String, ByVal sGameFolder As String) As Boolean

        Dim objFolder As Folder, objFSO As New FileSystemObject
        For Each objFolder In objFSO.GetFolder(sModsFolder).SubFolders
            objFSO.CopyFolder(objFolder.Path & "\*", sGameFolder, True)
        Next
        Return True
    End Function

    Private Function savePaths(ByVal sGameFolder As String, ByVal sModsFolder As String, ByVal sBackupFolder As String) As Boolean

        'Get %AppData% local folder
        Dim objFSO As New FileSystemObject, sPath As String
        sPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        If Not objFSO.FolderExists(sPath) Then
            Err.Raise(vbObjectError + 514, , "Invalid file path to AppData folder.")
            Return False
            Exit Function
        End If

        'Make DBG3MI subfolder
        sPath += "\DBG3MI\"
        If Not objFSO.FolderExists(sPath) Then
            objFSO.CreateFolder(sPath)
        End If

        'populate filename
        sPath += "FolderPaths.txt"

        'write the save file
        Using objWriter As New System.IO.StreamWriter(sPath)
            objWriter.WriteLine(sModsFolder)
            objWriter.WriteLine(sGameFolder)
            objWriter.WriteLine(sBackupFolder)
            objWriter.Flush()
            objWriter.Close()
        End Using

        Return True
    End Function

    'Returns an array of file paths less the root folder path within a folder, including all subfolders recursively.
    Public Function RecurseFolderFiles(ByRef r_objFolder As Folder, ByVal lPathLength As Long) As String()

        'Get all subfolder files
        Dim objFolder As Folder, objFSO As New FileSystemObject
        Dim sTempFiles() As String, sFiles As New List(Of String), lIndex As Long

        For Each objFolder In objFSO.GetFolder(r_objFolder.Path).SubFolders
            sTempFiles = RecurseFolderFiles(objFolder, lPathLength)
            For lIndex = LBound(sTempFiles) To UBound(sTempFiles)
                sFiles.Add(sTempFiles(lIndex))
            Next
        Next

        'Get folder files
        sTempFiles = ReturnFiles(r_objFolder, lPathLength)
        For lIndex = LBound(sTempFiles) To UBound(sTempFiles)
            sFiles.Add(sTempFiles(lIndex))
        Next
        Return sFiles.ToArray()
    End Function

    Public Function TestPaths() As List(Of String)

        'pull saved directories, or prompt for new ones
        Dim sPaths As New List(Of String)
        If txtGameFolder.Text <> vbNullString Then
            sPaths.Add(txtGameFolder.Text)
        Else
            sPaths.Add(GetFolderPath("Game"))
        End If
        If txtModsFolder.Text <> vbNullString Then
            sPaths.Add(txtModsFolder.Text)
        Else
            sPaths.Add(GetFolderPath("Mods"))
        End If
        If txtBackupFolder.Text <> vbNullString Then
            sPaths.Add(txtBackupFolder.Text)
        Else
            sPaths.Add(GetFolderPath("Backup"))
        End If

        'confirm folders exist
        Dim objFSO As New FileSystemObject, sPath As String
        For Each sPath In sPaths
            If Not objFSO.FolderExists(sPath) Then
                MsgBox("The below folder does not exist. Is the path correct?" + vbCrLf + sPath, vbOKOnly + vbExclamation, "Invalid Folder")
                sPaths.Remove(sPath)
            End If
        Next
        Return sPaths

    End Function

    'Returns an array of file paths less the root folder path within a given folder
    Public Shared Function ReturnFiles(ByRef r_objFolder As Folder, ByVal lPathLength As Long) As String()

        'Return all files in a given folder
        Dim objFile As File, sFiles As New List(Of String)
        For Each objFile In r_objFolder.Files
            sFiles.Add(objFile.Path.Substring(objFile.Path.Length - lPathLength - 1))
        Next objFile
        Return sFiles.ToArray()
    End Function

End Class
