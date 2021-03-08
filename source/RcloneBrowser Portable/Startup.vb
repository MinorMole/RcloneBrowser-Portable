Imports System.IO
Imports System.IO.Compression
Imports System.Windows.Forms
Imports Microsoft.Win32

Module Startup

    Dim Client As New Net.WebClient
    Dim UpdateStatus As Boolean = False
    Dim DeleteOldCount As Integer = 0
    Dim Retry_Count As Integer = 0

    Sub Main()

        '' Delete old version
        While My.Computer.FileSystem.FileExists(".\RcloneBrowser Portable.old") And DeleteOldCount < 10
            Try
                My.Computer.FileSystem.DeleteFile(".\RcloneBrowser Portable.old")
            Catch ex As Exception
                DeleteOldCount += 1
                Threading.Thread.Sleep(1000)
            End Try
        End While

        '' Check update
        Try
            Client.CachePolicy = New Net.Cache.RequestCachePolicy(Net.Cache.RequestCacheLevel.NoCacheNoStore)
            Dim CheckVersion As String = Client.DownloadString("https://download.konayuki.moe/RcloneBrowser-Portable/Version")
            If CheckVersion.Contains(".") And Application.ProductVersion.Trim <> CheckVersion Then Update(True)
        Catch ex As Exception
        End Try

        '' Check if RcloneBrowser and rclone exists
        If Not My.Computer.FileSystem.FileExists(".\RcloneBrowser\RcloneBrowser.exe") Or Not My.Computer.FileSystem.FileExists(".\RcloneBrowser\rclone.exe") Then Update(False)

        If UpdateStatus = True Then GoTo Quit

        '' Install WinFsp for RcloneBrowser Mounting Feature
        If Not My.Computer.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) + "\WinFsp") And Not My.Computer.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\WinFsp") Then
            Dim WinFsp As New ProcessStartInfo()
            WinFsp.WorkingDirectory = Application.StartupPath
            WinFsp.CreateNoWindow = True
            WinFsp.WindowStyle = ProcessWindowStyle.Hidden
            WinFsp.Verb = "runas"
            WinFsp.FileName = "cmd.exe"
            WinFsp.Arguments = "/c msiexec /i " + Chr(34) + Application.StartupPath + "\RcloneBrowser\redists\WinFsp.msi" + Chr(34) + " INSTALLLEVEL=1000 /quiet /qn /norestart"
            Try
                Process.Start(WinFsp)
            Catch ex As Exception
            End Try
        End If

        '' Check if rclone.conf exists
        If Not My.Computer.FileSystem.FileExists(".\RcloneBrowser\rclone.conf") Then
            File.Create(".\RcloneBrowser\rclone.conf").Dispose()
            MsgBox("Seem like you're using RcloneBrowser for the first time." + vbNewLine + vbNewLine + "Just let you know, rclone.conf is located at..." + vbNewLine + vbNewLine + Application.StartupPath + "\RcloneBrowser\rclone.conf")
            Process.Start("https://github.com/MinorMole/RcloneBrowser-Portable/wiki/RcloneBrowser-Guide")
            'Else
            '' Check if rclone.conf encrypt or not
            'Dim RcloneConf As String = My.Computer.FileSystem.ReadAllText(".\RcloneBrowser\rclone.conf")
            'If RcloneConf.Length.ToString <> 0 And Not RcloneConf.Contains("Encrypted rclone configuration File") And Not RcloneConf.Contains("RCLONE_ENCRYPT") Then MsgBox("Seem like you haven't created a password for RcloneBrowser!" + vbNewLine + vbNewLine + "To set a new password goto Config > Set configuration password")
        End If

        '' Check if the registry are correct
        Dim StartupPath As String = Application.StartupPath.Replace("\", "/")
        If (Registry.CurrentUser.OpenSubKey("Software\rclone-browser") Is Nothing) _
            Or (Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "rclone", "") <> StartupPath + "/RcloneBrowser/rclone.exe") _
            Or (Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "rcloneConf", "") <> StartupPath + "/RcloneBrowser/rclone.conf") Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "rclone", StartupPath + "/RcloneBrowser/rclone.exe")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "rcloneConf", StartupPath + "/RcloneBrowser/rclone.conf")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "alwaysShowInTray", "true")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "closeToTray", "true")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "notifyFinishedTransfers", "true")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "rowColors", "true")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "showFileIcons", "true")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "showFolderIcons", "true")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "showHidden", "true")
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Transfer", "checkVerbose", "true")
        End If

        '' Set Default Stream Command
        Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "stream", ".\RcloneBrowser\redists\mpv.exe - --title=" + Chr(34) + "$file_name" + Chr(34))
        Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "streamConfirmed", "true")

        '' Set Default Stream Command (Reverse)
        Dim ReadMountValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "mount", Nothing)
        If ReadMountValue = "--vfs-cache-mode writes --fast-list" Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "mount", "--vfs-cache-mode writes")
        End If

        '' Set Default Rclone Options
        Dim ReadRcloneOptions = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "defaultRcloneOptions", Nothing)
        If ReadRcloneOptions = Nothing Or ReadRcloneOptions = "--fast-list" Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "defaultRcloneOptions", "")
        End If

        '' Disable Auto Update of Rclone Browser
        Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "checkRcloneBrowserUpdates", "false")
        Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "checkRcloneUpdates", "false")

        '' Set Default Style
        Dim ReadiconsLayoutValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "iconsLayout", Nothing)
        If ReadiconsLayoutValue = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "iconsLayout", "longlist")
        End If
        Dim ReadiconSizeValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "iconSize", Nothing)
        If ReadiconsLayoutValue = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "iconSize", "S")
        End If
        Dim ReadbuttonSizeValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "buttonSize", Nothing)
        If ReadbuttonSizeValue = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "buttonSize", "0")
        End If
        Dim ReadfontSizeValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "fontSize", Nothing)
        If ReadfontSizeValue = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "fontSize", "0")
        End If
        Dim ReadiconsColour = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "iconsColour", Nothing)
        If ReadiconsColour = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "iconsColour", "black")
        End If
        Dim ReaddarkMode = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "darkMode", Nothing)
        If ReaddarkMode = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "darkMode", "false")
        End If
        Dim ReaddarkModeIni = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "darkModeIni", Nothing)
        If ReaddarkModeIni = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "darkModeIni", "false")
        End If
        Dim ReadbuttonSize = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "buttonSize", Nothing)
        If ReadbuttonSize = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "buttonSize", "0")
        End If
        Dim ReadbuttonStyle = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "buttonStyle", Nothing)
        If ReadbuttonStyle = Nothing Then
            Registry.SetValue("HKEY_CURRENT_USER\Software\rclone-browser\rclone-browser\Settings", "buttonStyle", "textandicon")
        End If

        '' Start RcloneBrowser
        If Process.GetProcessesByName("RcloneBrowser").Count = 0 Then
            Process.Start(".\RcloneBrowser\RcloneBrowser.exe")
            Try
                AppActivate("Rclone Browser")
            Catch ex As Exception
            End Try
        Else
            Try
                AppActivate("Rclone Browser")
            Catch ex As Exception
            End Try
            Process.Start(".\RcloneBrowser\RcloneBrowser.exe")

        End If

Quit:

    End Sub

    Function Update(ByRef ShowMsg As Boolean)

        UpdateStatus = True

        If ShowMsg Then MsgBox("A new version of RcloneBrowser Portable is avaliable!" + vbNewLine + vbNewLine + "We have to close RcloneBrowser and rclone to update" + vbNewLine + vbNewLine + "Click OK to update")

RTY_DL: Try
            My.Computer.Network.DownloadFile("https://download.konayuki.moe/RcloneBrowser-Portable/Release.zip", ".\Release.zip", vbNullString, vbNullString, True, 100000, True)
        Catch ex As Exception
            Retry_Count = Retry_Count + 1
            If Retry_Count > 10 Then
                MsgBox("Can't access to the update server, please try again later.")
                GoTo Quit
            Else
                Threading.Thread.Sleep(1000)
                GoTo RTY_DL
            End If
        End Try

        If My.Computer.FileSystem.DirectoryExists(".\tmp\") Then My.Computer.FileSystem.DeleteDirectory(".\tmp\", FileIO.DeleteDirectoryOption.DeleteAllContents)

        ZipFile.ExtractToDirectory(".\Release.zip", ".\tmp\")

        My.Computer.FileSystem.DeleteFile(".\Release.zip")

        My.Computer.FileSystem.MoveFile(Application.ExecutablePath, ".\RcloneBrowser Portable.old")

        Dim RcloneBrowserProcess = Process.GetProcessesByName("RcloneBrowser")
        For i As Integer = 0 To RcloneBrowserProcess.Count - 1
            RcloneBrowserProcess(i).Kill()
        Next i

        Dim rcloneProcess = Process.GetProcessesByName("rclone")
        For i As Integer = 0 To rcloneProcess.Count - 1
            rcloneProcess(i).Kill()
        Next i

        My.Computer.FileSystem.MoveDirectory(".\tmp\", ".\", overwrite:=True)

        Process.Start(".\RcloneBrowser Portable.exe")

Quit:   Application.Exit()

        Return Nothing

    End Function

End Module