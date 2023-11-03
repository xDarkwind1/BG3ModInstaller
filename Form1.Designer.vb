<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(frmMain))
        cmdInstall = New Button()
        cmdUninstall = New Button()
        txtModsFolder = New TextBox()
        lblModsFolder = New Label()
        lblGameDirectory = New Label()
        txtGameFolder = New TextBox()
        lblName = New Label()
        lblAuthor = New Label()
        lblBackupPath = New Label()
        txtBackupFolder = New TextBox()
        TextBox1 = New TextBox()
        cmdSetModPath = New Button()
        cmdSetGamePath = New Button()
        cmdSetBackupPath = New Button()
        cmdClose = New Button()
        SuspendLayout()
        ' 
        ' cmdInstall
        ' 
        cmdInstall.AccessibleName = "Install Mods"
        cmdInstall.Location = New Point(541, 275)
        cmdInstall.Name = "cmdInstall"
        cmdInstall.Size = New Size(100, 23)
        cmdInstall.TabIndex = 0
        cmdInstall.Text = "Install Mods"
        cmdInstall.UseVisualStyleBackColor = True
        ' 
        ' cmdUninstall
        ' 
        cmdUninstall.AccessibleName = "Uninstall Mods"
        cmdUninstall.Location = New Point(541, 318)
        cmdUninstall.Name = "cmdUninstall"
        cmdUninstall.Size = New Size(100, 23)
        cmdUninstall.TabIndex = 1
        cmdUninstall.Text = "Uninstall Mods"
        cmdUninstall.UseVisualStyleBackColor = True
        ' 
        ' txtModsFolder
        ' 
        txtModsFolder.Location = New Point(429, 109)
        txtModsFolder.Name = "txtModsFolder"
        txtModsFolder.Size = New Size(320, 23)
        txtModsFolder.TabIndex = 2
        ' 
        ' lblModsFolder
        ' 
        lblModsFolder.AutoSize = True
        lblModsFolder.Location = New Point(538, 91)
        lblModsFolder.Name = "lblModsFolder"
        lblModsFolder.Size = New Size(103, 15)
        lblModsFolder.TabIndex = 3
        lblModsFolder.Text = "Mods Folder Path:"
        ' 
        ' lblGameDirectory
        ' 
        lblGameDirectory.AutoSize = True
        lblGameDirectory.Location = New Point(537, 151)
        lblGameDirectory.Name = "lblGameDirectory"
        lblGameDirectory.Size = New Size(104, 15)
        lblGameDirectory.TabIndex = 5
        lblGameDirectory.Text = "Game Folder Path:"
        ' 
        ' txtGameFolder
        ' 
        txtGameFolder.Location = New Point(429, 169)
        txtGameFolder.Name = "txtGameFolder"
        txtGameFolder.Size = New Size(320, 23)
        txtGameFolder.TabIndex = 4
        ' 
        ' lblName
        ' 
        lblName.AutoSize = True
        lblName.Font = New Font("Segoe UI", 14.0F, FontStyle.Regular, GraphicsUnit.Point)
        lblName.Location = New Point(270, 36)
        lblName.Name = "lblName"
        lblName.Size = New Size(254, 25)
        lblName.TabIndex = 6
        lblName.Text = "Baldur's Gate 3 Mod Installer"
        ' 
        ' lblAuthor
        ' 
        lblAuthor.AutoSize = True
        lblAuthor.Font = New Font("Segoe UI", 8.0F, FontStyle.Regular, GraphicsUnit.Point)
        lblAuthor.Location = New Point(362, 23)
        lblAuthor.Name = "lblAuthor"
        lblAuthor.Size = New Size(65, 13)
        lblAuthor.TabIndex = 7
        lblAuthor.Text = "Darkwind's"
        ' 
        ' lblBackupPath
        ' 
        lblBackupPath.AutoSize = True
        lblBackupPath.Location = New Point(535, 211)
        lblBackupPath.Name = "lblBackupPath"
        lblBackupPath.Size = New Size(112, 15)
        lblBackupPath.TabIndex = 9
        lblBackupPath.Text = "Backup Folder Path:"
        ' 
        ' txtBackupFolder
        ' 
        txtBackupFolder.Location = New Point(429, 229)
        txtBackupFolder.Name = "txtBackupFolder"
        txtBackupFolder.Size = New Size(320, 23)
        txtBackupFolder.TabIndex = 8
        ' 
        ' TextBox1
        ' 
        TextBox1.Enabled = False
        TextBox1.Location = New Point(62, 103)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(348, 176)
        TextBox1.TabIndex = 10
        TextBox1.TabStop = False
        TextBox1.Text = resources.GetString("TextBox1.Text")
        ' 
        ' cmdSetModPath
        ' 
        cmdSetModPath.Image = My.Resources.Resource1.FileIcon
        cmdSetModPath.Location = New Point(755, 103)
        cmdSetModPath.Name = "cmdSetModPath"
        cmdSetModPath.Size = New Size(33, 33)
        cmdSetModPath.TabIndex = 11
        cmdSetModPath.UseVisualStyleBackColor = True
        ' 
        ' cmdSetGamePath
        ' 
        cmdSetGamePath.Image = My.Resources.Resource1.FileIcon
        cmdSetGamePath.Location = New Point(755, 163)
        cmdSetGamePath.Name = "cmdSetGamePath"
        cmdSetGamePath.Size = New Size(33, 33)
        cmdSetGamePath.TabIndex = 12
        cmdSetGamePath.UseVisualStyleBackColor = True
        ' 
        ' cmdSetBackupPath
        ' 
        cmdSetBackupPath.Image = My.Resources.Resource1.FileIcon
        cmdSetBackupPath.Location = New Point(755, 223)
        cmdSetBackupPath.Name = "cmdSetBackupPath"
        cmdSetBackupPath.Size = New Size(33, 33)
        cmdSetBackupPath.TabIndex = 13
        cmdSetBackupPath.UseVisualStyleBackColor = True
        ' 
        ' cmdClose
        ' 
        cmdClose.Location = New Point(176, 306)
        cmdClose.Name = "cmdClose"
        cmdClose.Size = New Size(101, 23)
        cmdClose.TabIndex = 14
        cmdClose.Text = "Close Installer"
        cmdClose.UseVisualStyleBackColor = True
        ' 
        ' frmMain
        ' 
        AccessibleName = "Main"
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 366)
        Controls.Add(cmdClose)
        Controls.Add(cmdSetBackupPath)
        Controls.Add(cmdSetGamePath)
        Controls.Add(cmdSetModPath)
        Controls.Add(TextBox1)
        Controls.Add(lblBackupPath)
        Controls.Add(txtBackupFolder)
        Controls.Add(lblAuthor)
        Controls.Add(lblName)
        Controls.Add(lblGameDirectory)
        Controls.Add(txtGameFolder)
        Controls.Add(lblModsFolder)
        Controls.Add(txtModsFolder)
        Controls.Add(cmdUninstall)
        Controls.Add(cmdInstall)
        Name = "frmMain"
        Text = "Darkwind's Baldur's Gate 3 Mod Installer"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents cmdInstall As Button
    Friend WithEvents cmdUninstall As Button
    Friend WithEvents txtModsFolder As TextBox
    Friend WithEvents lblModsFolder As Label
    Friend WithEvents lblGameDirectory As Label
    Friend WithEvents txtGameFolder As TextBox
    Friend WithEvents lblName As Label
    Friend WithEvents lblAuthor As Label
    Friend WithEvents lblBackupPath As Label
    Friend WithEvents txtBackupFolder As TextBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents cmdSetModPath As Button
    Friend WithEvents cmdSetGamePath As Button
    Friend WithEvents cmdSetBackupPath As Button
    Friend WithEvents cmdClose As Button
End Class
