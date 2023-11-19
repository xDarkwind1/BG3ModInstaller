Option Explicit On
Imports Scripting

Public Class frmMain

    'Loads settings file on open
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo PROC_ERR

        'Get appdata folder if it exists
        Dim objFSO As New FileSystemObject, sPath As String
        sPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\DBG3MI\"
        If Not objFSO.FolderExists(sPath) Then
            objFSO.CreateFolder(sPath)
        End If

        'read the FolderPaths text file and push to textboxes
        Call readPathsFile(sPath)

PROC_ERR:
        Exit Sub
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
            txtGameFolder.Text = sPath
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
        Dim sFiles As New List(Of String)
        Dim objFSO As New FileSystemObject, objFolder As Folder
        For Each objFolder In objFSO.GetFolder(sModsFolder).SubFolders
            sFiles.AddRange(getInstallationFiles(objFolder.Path, objFolder.Path.Length + 1).ToList)
        Next

        'Exit if no files to install and warn the user
        If sFiles.Count = 0 Then
            MsgBox("No files found in the supplied directory. Is the mods folder correct?", vbOKOnly + vbExclamation, "No File Found")
            Exit Sub
        End If

        'Saves list of installation files
        If Not saveInstallationFiles(sBackupFolder, sFiles.ToArray) Then
            MsgBox("Could not save the list of installation files to the Backup Folder. Is the \InstalledFiles.txt file open?", vbOKOnly + vbExclamation, "Could Not Save File List")
            Exit Sub
        End If

        'Save copy of files that will be overwritten
        If Not saveOriginalFiles(sGameFolder, sBackupFolder, sFiles.ToArray) Then
            MsgBox("Could not copy existing game files to the Backup Folder. Do you have write permissions for that folder?", vbOKOnly + vbExclamation, "Could Not Backup Files")
            Exit Sub
        End If

        'Copy over mod files to directory
        If Not copyModFiles(sModsFolder, sGameFolder) Then
            MsgBox("Failed to copy mod files into the game folder. Installation may have begun before this error occurred! Check the game folder")
            Exit Sub
        End If

        'Save the used paths to appdata
        If Not savePaths(sGameFolder, sModsFolder, sBackupFolder) Then
            If MsgBox("Failed to save folder paths for future use. Installation otherwise completed successfully. Close the installer?", vbYesNo + vbQuestion, "Close Installer?") = vbYes Then
                Application.Exit()
            End If
            Exit Sub
        End If

        'Inform of success and provide option to exit
        If MsgBox("Installation completed successfully! Close the installer?", vbYesNo + vbQuestion, "Close Installer?") = vbYes Then
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
        Dim sFiles As List(Of String) = readInstallationFiles(sBackupFolder)

        'Remove mod files from game folder
        If Not DeleteModFiles(sGameFolder, sFiles) Then
            MsgBox("At least one file from the installation definition is missing from the directory. Uninstallation cannot be completed. Recommend a clean install of BG3.", vbOKOnly + vbExclamation)
            Exit Sub
        End If

        'Restore original files
        If Not restoreGameFiles(sGameFolder, sBackupFolder) Then
            MsgBox("At least one file from the Backup Folder could not be restored. Recommend a clean install of BG3.", vbOKOnly + vbExclamation)
            Exit Sub
        End If

        'Inform of success and provide option to exit
        If MsgBox("Uninstalled successfully. Close the installer?", vbYesNo + vbQuestion, "Close Installer?") = vbYes Then
            Application.Exit()
        End If

    End Sub

    'Closes the installer
    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Application.Exit()
    End Sub

    'Populates textboxes with data from saved Paths file.
    Private Sub readPathsFile(ByVal sPath As String)
        On Error GoTo PROC_ERR

        Dim objFSO As New FileSystemObject
        If Not objFSO.FolderExists(sPath) Then
            Err.Raise(vbObjectError + 514, , "Invalid file path argument when reading the Paths file.")
            Exit Sub
        End If

        If objFSO.FileExists(sPath & "FolderPaths.txt") Then
            Dim objReader As New System.IO.StreamReader(sPath & "FolderPaths.txt")
            Dim sTextLine As String, lIndex As Long
            For lIndex = 1 To 3
                If objReader.Peek() = -1 Then Exit For
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
            objReader.Close()
        End If
PROC_ERR:
        Exit Sub
    End Sub

    'Returns a folder path based on folderdialog
    Public Shared Function GetFolderPath(ByVal sPathType As String) As String
        Dim sPath As String, objFB As New myFolderBrowser With {
            .Description = "Select the " & sPathType & " folder...",
            .StartLocation = myFolderBrowser.enuFolderBrowserFolder.MyComputer
        }
        If objFB.ShowBrowser = DialogResult.OK Then
            sPath = objFB.Path
        Else
            sPath = String.Empty
        End If
        Return sPath
    End Function

    'Returns an array of file path names relative to sModsFolder containing each file in the subdirectories
    Private Function getInstallationFiles(ByVal sModFolder As String, ByVal lPathLength As Long) As String()
        On Error GoTo PROC_ERR

        'Check nonexistant folder
        Dim objFSO As New FileSystemObject
        If Not objFSO.FolderExists(sModFolder) Then
            Return Split("")
        End If

        'recurse files in folder
        Dim objFolder As Folder
        objFolder = objFSO.GetFolder(sModFolder)
        Return RecurseFolderFiles(objFolder, lPathLength)
PROC_ERR:
        Return Split("")
    End Function

    'Saves a text list of each file installed
    Private Function saveInstallationFiles(ByVal sBackupFolder As String, ByVal sFiles() As String) As Boolean
        On Error GoTo PROC_ERR
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
PROC_ERR:
        Return False
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
        On Error GoTo PROC_ERR
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
            On Error Resume Next
            objFSO.CopyFile(sGameFolder + sFile, sDir, True)
            On Error GoTo PROC_ERR
        Next
        Return True
PROC_ERR:
        Return False
    End Function

    'Copies mod files into game folder
    Private Function copyModFiles(ByVal sModsFolder As String, ByVal sGameFolder As String) As Boolean
        On Error GoTo PROC_ERR
        Dim objFolder As Folder, objFSO As New FileSystemObject
        For Each objFolder In objFSO.GetFolder(sModsFolder).SubFolders
            objFSO.CopyFolder(objFolder.Path & "\*", sGameFolder, True)
        Next
        Return True
PROC_ERR:
        Return False
    End Function

    Private Function restoreGameFiles(ByVal sGameFolder As String, ByVal sBackupFolder As String) As Boolean
        On Error GoTo PROC_ERR
        Dim ObjFSO As New FileSystemObject
        ObjFSO.CopyFolder(sBackupFolder + "\*", sGameFolder + "\", True)
        Return True
PROC_EXIT:
        Exit Function

PROC_ERR:
        Return False
    End Function

    Private Function savePaths(ByVal sGameFolder As String, ByVal sModsFolder As String, ByVal sBackupFolder As String) As Boolean
        On Error GoTo PROC_ERR

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
PROC_ERR:
        Return False
    End Function

    'Returns an array of file paths less the root folder path within a folder, including all subfolders recursively.
    Public Function RecurseFolderFiles(ByRef r_objFolder As Folder, ByVal lPathLength As Long) As String()
        On Error GoTo PROC_ERR
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
PROC_ERR:
        Return Split("")
    End Function

    Public Function TestPaths() As List(Of String)
        On Error GoTo PROC_ERR
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
PROC_ERR:
        sPaths.Clear()
        Return sPaths
    End Function

    'Deletes mod files in prep of return to vanilla state
    Public Function DeleteModFiles(ByVal sGameFolder As String, ByVal sFiles As List(Of String)) As Boolean
        On Error GoTo PROC_ERR

        'loop thru all files in the definition
        Dim sFile As String, objFSO As New FileSystemObject
        If sFiles.Count < 1 Then Return True
        For Each sFile In sFiles

            'build full file path
            Dim sPath As String
            sPath = sGameFolder + "\" + sFile

            'Confirm file exists in the game folder, then delete if present
            If objFSO.FileExists(sGameFolder + "\" + sFile) Then
                objFSO.DeleteFile(sGameFolder + "\" + sFile)

                'return a failure if a file doesn't exist
            Else Return False
            End If

            'Remove each part of the path segment by segment
            Do Until sPath.Length <= sGameFolder.Length
                sPath = sPath.Remove(sPath.LastIndexOf("\"))

                'test if anything is in each folder. if not, delete it
                If objFSO.GetFolder(sPath).Files.Count = 0 And objFSO.GetFolder(sPath).SubFolders.Count = 0 Then
                    objFSO.DeleteFolder(sPath)
                Else
                    Exit Do
                End If
            Loop
        Next
        Return True

PROC_ERR:
        Return False
    End Function

    'Returns an array of file paths less the root folder path within a given folder
    Public Shared Function ReturnFiles(ByRef r_objFolder As Folder, ByVal lPathLength As Long) As String()
        On Error GoTo PROC_ERR
        'Return all files in a given folder
        Dim objFile As File, sFiles As New List(Of String)
        For Each objFile In r_objFolder.Files
            sFiles.Add(objFile.Path.Substring(lPathLength))
        Next objFile
        Return sFiles.ToArray()
PROC_ERR:
        Return Split("")
    End Function

End Class
