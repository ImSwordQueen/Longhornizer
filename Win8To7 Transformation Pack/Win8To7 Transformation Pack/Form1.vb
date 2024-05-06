Public Class Form1
#Region "Variables"
    Private windrive = IO.Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System))
    Private windir As String = System.Environment.GetEnvironmentVariable("WINDIR")
    Private storagelocation = windir + "\Longhornizer"
    Private sysdir As String
    Public HKLMKey32 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine

    Public isInstalled As Boolean = False 'has already been installed?
#End Region
#Region "HC compatibility"
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Const WM_THEMECHANGED = &H31A

        If (m.Msg = WM_THEMECHANGED) Then
            RefreshHC()
            Me.Refresh()
        End If

        MyBase.WndProc(m)
    End Sub

    Public Sub RefreshHC()
        If SystemInformation.HighContrast = True Then
            Me.BackColor = Control.DefaultBackColor
            Me.BackgroundImage = Nothing
            Me.ForeColor = Control.DefaultForeColor
            Next1.FlatStyle = FlatStyle.System
            LinkLabel1.ForeColor = Control.DefaultForeColor
            LinkLabel1.LinkColor = Control.DefaultForeColor
            LinkLabel1.ActiveLinkColor = Control.DefaultForeColor
            LinkLabel1.VisitedLinkColor = Control.DefaultForeColor
            LinkLabel2.ForeColor = Control.DefaultForeColor
            LinkLabel2.LinkColor = Control.DefaultForeColor
            LinkLabel2.ActiveLinkColor = Control.DefaultForeColor
            LinkLabel2.VisitedLinkColor = Control.DefaultForeColor
        Else
            Me.BackColor = Color.Black
            Me.BackgroundImage = My.Resources.int_SetupBG
            Me.ForeColor = Color.White
            Next1.FlatStyle = FlatStyle.Flat
            LinkLabel1.ForeColor = Color.White
            LinkLabel1.LinkColor = Color.White
            LinkLabel1.ActiveLinkColor = Color.White
            LinkLabel1.VisitedLinkColor = Color.White
            LinkLabel2.ForeColor = Color.White
            LinkLabel2.LinkColor = Color.White
            LinkLabel2.ActiveLinkColor = Color.White
            LinkLabel2.VisitedLinkColor = Color.White
        End If
    End Sub
#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Check System32 directory
        If System.IO.File.Exists(windir + "\SysNative\LogonUI.exe") Then
            sysdir = "SysNative"
        Else
            sysdir = "System32"
        End If
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If PageTransform.Visible = True Then
            MsgBox("Installation of Longhornizer Transformation Pack  cannot be cancelled, as cancelling the transformation can risk rendering Windows unusable.", MsgBoxStyle.Exclamation, "Longhornizer Transformation Pack ")
            e.Cancel = True
        End If
    End Sub

#Region "Label status changes"
    Private Sub AboutToRestart(ByVal status As String)
        'This bit here makes it work despite being called via a VB.NET Thread
        If Me.InvokeRequired Then
            Dim args() As String = {status}
            Me.Invoke(New Action(Of String)(AddressOf AboutToRestart), args)
            Exit Sub
        End If

        ProgressBar1.Style = ProgressBarStyle.Marquee
        LabelStatus.Text = "Windows will restart in a few moments to continue the transformation process..."
    End Sub

    Private Sub RestartTime(ByVal status As String)
        'This bit here makes it work despite being called via a VB.NET Thread
        If Me.InvokeRequired Then
            Dim args() As String = {status}
            Me.Invoke(New Action(Of String)(AddressOf RestartTime), args)
            Exit Sub
        End If

        PageTransform.Visible = False
        End
    End Sub

    Private Sub ChangeProgress(ByVal progress As Integer)
        'This bit here makes it work despite being called via a VB.NET Thread
        If Me.InvokeRequired Then
            Dim args() As String = {progress}
            Me.Invoke(New Action(Of String)(AddressOf ChangeProgress), args)
            Exit Sub
        End If

        ProgressBar1.Value = progress
    End Sub

    Private Sub ChangeProgressStyle(ByVal style As ProgressBarStyle)
        'This bit here makes it work despite being called via a VB.NET Thread
        If Me.InvokeRequired Then
            Dim args() As String = {style}
            Me.Invoke(New Action(Of String)(AddressOf ChangeProgressStyle), args)
            Return
        End If

        ProgressBar1.Style = style
    End Sub

    Private Sub ChangeStatus(ByVal status As String)
        'This bit here makes it work despite being called via a VB.NET Thread
        If Me.InvokeRequired Then
            Dim args() As String = {status}
            Me.Invoke(New Action(Of String)(AddressOf ChangeStatus), args)
            Exit Sub
        End If

        LabelStatus.Text = "Status: " + status
    End Sub
#End Region

#Region "First page - new installs"
    Private Sub Next1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Next1.Click
        BackArea.Visible = True
        BackButton.Enabled = True
        Page2.Visible = True
        Page1.Visible = False
    End Sub
    Private Sub Next1_Hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Next1.MouseEnter, Next1.MouseUp
        Next1.BackgroundImage = My.Resources.int_ButtonHover
        Next1.Image = My.Resources.int_ForwardHover
    End Sub
    Private Sub Next1_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Next1.MouseLeave
        Next1.BackgroundImage = Nothing
        Next1.Image = My.Resources.int_ForwardNormal
    End Sub
    Private Sub Next1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Next1.MouseDown
        Next1.BackgroundImage = My.Resources.int_ButtonPressed
        Next1.Image = My.Resources.int_ForwardPressed
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Credits.Show()
        Credits.Focus()
    End Sub
#End Region
#Region "First page - update mode"
    Private Sub ManageOptionConfigure_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManageOptionConfigure.CheckedChanged
        If ManageOptionConfigure.Checked = True Then
            NextManage.Text = "Next"
        Else
            NextManage.Text = "Launch Uninstaller"
        End If
    End Sub

    Private Sub NextManage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextManage.Click
        If ManageOptionConfigure.Checked = True Then
            Page2.Visible = True
            PageOptions.Visible = False
            BackButton.Enabled = True
        Else
            Shell(storagelocation + "\SetupTools\setup.exe", AppWinStyle.NormalFocus, False) 'Go over to uninstaller
            End
        End If
    End Sub
#End Region
#Region "Second page - Ts&Cs"
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Next2.Enabled = CheckBox1.Checked
    End Sub

    Private Sub Next2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Next2.Click
        PageConfirm.Visible = True
        Page2.Visible = False
        BackButton.Enabled = True
    End Sub
#End Region
#Region "Confirmation page"
    Private Sub NextConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextConfirm.Click

        BackArea.Visible = False
        PageTransform.Visible = True
        PageConfirm.Visible = False

        Dim jobthread As New Thread(AddressOf DoTheJob)
        jobthread.Start()
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllowUXThemePatcher.CheckedChanged
        If AllowUXThemePatcher.Checked = False Then
            If MsgBox("Not patching your system with UXThemePatcher SHOULD ONLY BE DONE if you have already patched your system to be able to use custom themes already. If not patched, the transformation will brick your current Windows installation. Are you SURE you want to continue without patching your system with UXThemePatcher?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Longhornizer Transformation Pack ") = MsgBoxResult.No Then
                AllowUXThemePatcher.Checked = True
            End If
        End If
    End Sub
#End Region

#Region "Important calls"
    Private Sub ErrorOccurred(ByVal status As String)
        'This bit here makes it work despite being called via a VB.NET Thread
        If Me.InvokeRequired Then
            Dim args() As String = {status}
            Me.Invoke(New Action(Of String)(AddressOf ErrorOccurred), args)
            Exit Sub
        End If

        If status = "" Then
            MsgBox("Something happened", "Something happened")
        Else
            MsgBox(status, MsgBoxStyle.OkOnly, "Something happened")
        End If
        End
    End Sub
#End Region

#Region "Back Button"
    Private Sub BackButton_EnabledChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackButton.EnabledChanged
        If BackButton.Enabled = True Then
            BackButton.Image = My.Resources.int_BackNormal
        Else
            BackButton.Image = My.Resources.int_BackDisabled
        End If
    End Sub
    Private Sub BackButton_Hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackButton.MouseEnter, BackButton.MouseUp
        If BackButton.Enabled = True Then
            BackButton.Image = My.Resources.int_BackHover
        End If
    End Sub
    Private Sub BackButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackButton.MouseLeave
        If BackButton.Enabled = True Then
            BackButton.Image = My.Resources.int_BackNormal
        End If
    End Sub
    Private Sub BackButton_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BackButton.MouseDown
        If BackButton.Enabled = True Then
            BackButton.Image = My.Resources.int_BackPressed
        End If
    End Sub

    Private Sub BackButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackButton.Click
        'Page 2 -> 1
        If Page2.Visible = True Then
            Page1.Visible = True
            Page2.Visible = False
            BackArea.Visible = False
        End If
        'Page 3 -> 2
        If PageConfirm.Visible = True Then
            If isInstalled = True Then
                PageOptions.Visible = True
                BackButton.Enabled = False
            Else
                Page2.Visible = True
            End If
            Page2.Visible = False
        End If
        'Page Confirm -> 5
        If PageConfirm.Visible = True Then
            Page2.Visible = True
            PageConfirm.Visible = False
        End If
    End Sub
#End Region
#Region "Credits button"
    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Credits.Show()
    End Sub
#End Region

#Region "Transformation begins"
    Sub DoTheJob()
        'Get sysprefix value first
        Dim sysprefix As String
        If System.IO.File.Exists(windir + "\SysNative\LogonUI.exe") Then
            sysprefix = windir + "\SysNative"
        Else
            sysprefix = windir + "\System32"
        End If

        'Create a Restore Point if this is the first installation
        If isInstalled = False Then
            ChangeStatus("Creating a restore point...")
            ChangeProgress(1)
            ChangeProgressStyle(ProgressBarStyle.Marquee)
            Shell(sysprefix + "\cmd.exe /c reg add ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"" /v ""SystemRestorePointCreationFrequency"" /t REG_DWORD /d 0 /f", AppWinStyle.Hide, True)
            Shell(sysprefix + "\cmd.exe /c wmic.exe /Namespace:\\root\default Path SystemRestore Call CreateRestorePoint ""Before installing Longhornizer Transformation Pack "", 100, 12", AppWinStyle.Hide, True)
            Shell(sysprefix + "\cmd.exe /c reg del ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"" /v ""SystemRestorePointCreationFrequency"" /f", AppWinStyle.Hide, True)
            ChangeProgressStyle(ProgressBarStyle.Continuous)

            '...and create the pack's directory
            ChangeStatus("Creating directory for transformation pack files...")
            ChangeProgress(2)
            'First, make the location a thing
            If IO.Directory.Exists(storagelocation) Then
                Try
                    IO.Directory.Delete(storagelocation, True)
                Catch ex As Exception
                    ErrorOccurred("An error occurred clearing room for the transformation pack directory.")
                    Exit Sub
                End Try
            End If
            If Not IO.Directory.Exists(storagelocation) Then
                Try
                    IO.Directory.CreateDirectory(storagelocation)
                Catch ex As Exception
                    ErrorOccurred(ex.Message)
                    Exit Sub
                End Try
            End If
        End If

        'Now extract the files
        ChangeStatus("Extracting files ready for patching...")
        ChangeProgress(3)

        If isInstalled = True Then
            Try 'Move the old transformation pack's version of Setup Mode Phases to a temporary spot
                IO.File.Move(storagelocation + "\SetupTools\setup.exe", storagelocation + "\setupold.exe")
            Catch ex As Exception
                ErrorOccurred("Couldn't move the old version of Setup for usage in Stage 2 of customisation")
                Exit Sub
            End Try
        End If

        If isInstalled = True And IO.File.Exists(storagelocation + "\SetupFiles\regchange.exe") Then '3.0 update-path support 1/2
            Try 'Move the old transformation pack's version of Registry Changes Executable to a temporary spot
                IO.File.Move(storagelocation + "\SetupFiles\regchange.exe", storagelocation + "\regchangeold.exe")
            Catch ex As Exception
                ErrorOccurred("Couldn't move the old version of Registry Changes Executable for usage in Stage 2 of customisation")
                Exit Sub
            End Try
        End If

        Dim tries As Integer 'Delete and recreate directories
        For Each direc In {storagelocation + "\FileReplacements", storagelocation + "\UserFileReplacements", storagelocation + "\ResFiles", storagelocation + "\SetupFiles", windir + "\Temp\LonghornizerDiscardedFiles"}
            tries = 0
            While Not tries = 10
                Shell(sysprefix + "\cmd.exe /c del """ + direc + """ /s /q /f /a", AppWinStyle.Hide, True)
                Shell(sysprefix + "\cmd.exe /c rd """ + direc + """ /s /q", AppWinStyle.Hide, True)
                If Not IO.Directory.Exists(direc) Then
                    Exit While
                End If
                tries += 1
            End While
        Next
        For Each direc In {storagelocation + "\Backups", storagelocation + "\FileReplacements", storagelocation + "\UserFileReplacements", storagelocation + "\ResFiles", storagelocation + "\SetupTools", _
                           windir + "\Temp\LonghornizerDiscardedFiles", storagelocation + "\SetupFiles"}
            If Not IO.Directory.Exists(direc) Then
                Try
                    IO.Directory.CreateDirectory(direc)
                Catch ex As Exception
                    ErrorOccurred(ex.Message)
                    Exit Sub
                End Try
            End If
        Next

        If isInstalled = True And IO.File.Exists(storagelocation + "\regchangeold.exe") Then '3.0 update-path support 2/2
            Try 'Move the old transformation pack's version of Registry Changes Executable to a temporary spot
                IO.File.Move(storagelocation + "\regchangeold.exe", storagelocation + "\SetupFiles\regchange.exe")
            Catch ex As Exception
                ErrorOccurred("Couldn't move the old version of Registry Changes Executable back to original placement for usage in Stage 2 of customisation")
                Exit Sub
            End Try
        End If

        Dim targetPath As String
        Dim tempArray As String()

        'Extract files from resources
        Dim startInfo As New System.Diagnostics.ProcessStartInfo
        Dim MyProcess As Process

        'First run: Only setupfiles and setuptools
        For Each dictEntry In My.Resources.ResourceManager.GetResourceSet(System.Globalization.CultureInfo.CurrentCulture, True, True)
            If dictEntry.Key.StartsWith("int_") Then 'Skip internal resources
                Continue For
            End If
            If dictEntry.Key.StartsWith("i386_") And Environment.Is64BitOperatingSystem = True Then 'Skip incompatible architectures - 64-Bit
                Continue For
            End If
            If dictEntry.Key.StartsWith("amd64_") And Environment.Is64BitOperatingSystem = False Then 'Skip incompatible architectures - 32-Bit
                Continue For
            End If

            ' Strip architecture identifier from id for next steps
            tempArray = dictEntry.Key.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)

            ' Strip build identifier from id for next steps
            tempArray = targetPath.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)

            ' Strip SKU identifier from id for next steps
            tempArray = targetPath.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)

            ' Now skip depending on remaining preferences
            If targetPath.StartsWith("AllowUXThemePatcher_") And AllowUXThemePatcher.Checked = False Then 'Skip UXThemePatcher stuff if not chosen
                Continue For
            End If

            ' Strip setting identifier from id for next steps
            tempArray = targetPath.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)


            'This is where you'd skip depending on Aesthetics not being selected, but no other aesthetic options are pre-coded in this, so... this part is empty.

            ' Strip aesthetic identifier from id for next steps
            tempArray = targetPath.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)


            'First, deal with Setup files
            If targetPath.StartsWith("setupfile:") Then
                'Create directory for extraction
                targetPath = targetPath.Replace("setupfile:", "")

                'Now copy the new file into the temporary path
                If Not IO.File.Exists(storagelocation + "\SetupFiles\" + targetPath) Then
                    If WriteFileFromResources(dictEntry.Key.ToString(), storagelocation + "\SetupFiles\" + targetPath) = False Then
                        ErrorOccurred("Failed to extract files")
                        Exit Sub
                    End If
                End If
            End If

            'Second, deal with Setup tools
            If targetPath.StartsWith("setuptool:") Then
                'Create directory for extraction
                targetPath = targetPath.Replace("setuptool:", "")

                'Now copy the new file into the temporary path
                If Not IO.File.Exists(storagelocation + "\SetupTools\" + targetPath) Then
                    If WriteFileFromResources(dictEntry.Key.ToString(), storagelocation + "\SetupTools\" + targetPath) = False Then
                        ErrorOccurred("Failed to extract files")
                        Exit Sub
                    End If
                End If
            End If
        Next

        'Second run: Remaining files
        For Each dictEntry In My.Resources.ResourceManager.GetResourceSet(System.Globalization.CultureInfo.CurrentCulture, True, True)
            If dictEntry.Key.StartsWith("int_") Then 'Skip internal resources
                Continue For
            End If
            If dictEntry.Key.StartsWith("i386_") And Environment.Is64BitOperatingSystem = True Then 'Skip incompatible architectures - 64-Bit
                Continue For
            End If
            If dictEntry.Key.StartsWith("amd64_") And Environment.Is64BitOperatingSystem = False Then 'Skip incompatible architectures - 32-Bit
                Continue For
            End If

            ' Strip architecture identifier from id for next steps
            tempArray = dictEntry.Key.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)

            ' Strip build identifier from id for next steps
            tempArray = targetPath.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)

            ' Strip SKU identifier from id for next steps
            tempArray = targetPath.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)

            ' Now skip depending on remaining preferences
            If targetPath.StartsWith("AllowUXThemePatcher_") And AllowUXThemePatcher.Checked = False Then 'Skip UXThemePatcher stuff if not chosen
                Continue For
            End If

            ' Strip setting identifier from id for next steps
            tempArray = targetPath.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)


            'This is where you'd skip depending on Aesthetics not being selected, but no other aesthetic options are pre-coded in this, so... this part is empty.

            ' Strip aesthetic identifier from id for next steps
            tempArray = targetPath.Split("_")
            tempArray = tempArray.Skip(1).ToArray()
            targetPath = String.Join("_", tempArray.ToArray)


            'Skip extracted Setup files and tools
            If targetPath.StartsWith("setupfile:") Or targetPath.StartsWith("setuptool:") Then
                Continue For
            End If

            'Now, get on with extracting these archives
            'First, deal with ones that are just straight-up file replacements
            If targetPath = "location" Then
                If WriteFileFromResources(dictEntry.Key.ToString(), storagelocation + "\TEMPfiles.7z") = False Then
                    ErrorOccurred("Failed to extract files")
                    Exit Sub
                End If

                startInfo = New System.Diagnostics.ProcessStartInfo
                startInfo.FileName = storagelocation + "\SetupTools\7za.exe"
                startInfo.Arguments = "x " + storagelocation + "\TEMPfiles.7z -o" + storagelocation + "\FileReplacements * -r"
                startInfo.CreateNoWindow = True
                startInfo.UseShellExecute = True
                startInfo.WindowStyle = ProcessWindowStyle.Hidden

                MyProcess = Process.Start(startInfo)
                MyProcess.WaitForExit()
                IO.File.Delete(storagelocation + "\TEMPfiles.7z")
                If (MyProcess.ExitCode <> 0) Then
                    ErrorOccurred("Failed to extract files")
                    Exit Sub
                End If
            End If

            'Second, deal with ones that are straight-up file replacements... for users.
            If targetPath = "userlocation" Then
                If WriteFileFromResources(dictEntry.Key.ToString(), storagelocation + "\TEMPuserfiles.7z") = False Then
                    ErrorOccurred("Failed to extract files")
                    Exit Sub
                End If

                startInfo = New System.Diagnostics.ProcessStartInfo
                startInfo.FileName = storagelocation + "\SetupTools\7za.exe"
                startInfo.Arguments = "x " + storagelocation + "\TEMPuserfiles.7z -o" + storagelocation + "\UserFileReplacements * -r"
                startInfo.CreateNoWindow = True
                startInfo.UseShellExecute = True
                startInfo.WindowStyle = ProcessWindowStyle.Hidden

                MyProcess = Process.Start(startInfo)
                MyProcess.WaitForExit()
                IO.File.Delete(storagelocation + "\TEMPuserfiles.7z")
                If (MyProcess.ExitCode <> 0) Then
                    ErrorOccurred("Failed to extract files")
                    Exit Sub
                End If
            End If

            'Second, deal with ones that are .res files
            If targetPath = "res" Then
                If WriteFileFromResources(dictEntry.Key.ToString(), storagelocation + "\TEMPres.7z") = False Then
                    ErrorOccurred("Failed to extract files")
                    Exit Sub
                End If

                startInfo = New System.Diagnostics.ProcessStartInfo
                startInfo.FileName = storagelocation + "\SetupTools\7za.exe"
                startInfo.Arguments = "x " + storagelocation + "\TEMPres.7z -o" + storagelocation + "\ResFiles * -r"
                startInfo.CreateNoWindow = True
                startInfo.UseShellExecute = True
                startInfo.WindowStyle = ProcessWindowStyle.Hidden

                MyProcess = Process.Start(startInfo)
                MyProcess.WaitForExit()
                IO.File.Delete(storagelocation + "\TEMPres.7z")
                If (MyProcess.ExitCode <> 0) Then
                    ErrorOccurred("Failed to extract files")
                    Exit Sub
                End If
            End If
        Next

        'Replace percentage paths with their actual paths

        'Clean up 7z files - it's no longer needed
        IO.File.Delete(storagelocation + "\SetupTools\7z License.txt")
        IO.File.Delete(storagelocation + "\SetupTools\7zxa.dll")
        IO.File.Delete(storagelocation + "\SetupTools\7za.dll")
        IO.File.Delete(storagelocation + "\SetupTools\7za.exe")

        ChangeProgress(4)
        'Configure the Registry options
        ChangeStatus("Configuring Registry values...")

        HKLMKey32.CreateSubKey("SOFTWARE\Longhornizer")
        'Remaining preferences
        If AllowUXThemePatcher.Checked = True Then
            HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer", True).SetValue("AllowUXThemePatcher", "true")
        Else
            HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer", True).SetValue("AllowUXThemePatcher", "false")
        End If

        'Remove transformation failure indicator if the transformation failed last time
        HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer", True).DeleteValue("TransformationFailed", False)

        If isInstalled = False Then
            'Set Phase to 2
            HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer", True).SetValue("CurrentPhase", 2)

            ChangeProgress(5)
            ChangeProgressStyle(ProgressBarStyle.Marquee)
            ChangeStatus("Windows will begin transformation in a few moments...")
            Thread.Sleep(8000)

            Shell("taskkill /f /im explorer.exe", AppWinStyle.Hide, False)
            Shell(storagelocation + "\SetupTools\setup.exe", AppWinStyle.MaximizedFocus, False)
            RestartTime("")
        Else
            'Set Phase to 102 (Configuration Phase 2)
            HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer", True).SetValue("CurrentPhase", 102)

            ChangeProgress(5)
            ChangeProgressStyle(ProgressBarStyle.Marquee)
            ChangeStatus("Windows will restart in a few moments...")
            Thread.Sleep(8000)

            'Switch to progress form
            If SystemInformation.HighContrast = True Then
                HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer", True).SetValue("HighContrast", 1)
            Else
                HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer", True).SetValue("HighContrast", 0)
            End If
            Shell(sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v CmdLine /d " + storagelocation + "\setupold.exe" + " /t REG_SZ /f", AppWinStyle.Hide, True)
            Shell(sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v OOBEInProgress /t REG_DWORD /d 1 /f", AppWinStyle.Hide, True)
            Shell(sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v SetupPhase /t REG_DWORD /d 4 /f", AppWinStyle.Hide, True)
            Shell(sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v SetupType /t REG_DWORD /d 2 /f", AppWinStyle.Hide, True)
            Shell(sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v SystemSetupInProgress /t REG_DWORD /d 1 /f", AppWinStyle.Hide, True)
            Shell(sysprefix + "\cmd.exe /c bcdedit /deletevalue {current} safeboot", AppWinStyle.Hide, True)
            Shell(sysprefix + "\cmd.exe /c shutdown /r /t 0 /f", AppWinStyle.Hide, False)
            RestartTime("")
        End If
    End Sub

    Function WriteFileFromResources(ByVal resourceID As String, ByVal targetPath As String)
        Dim audioStream As IO.MemoryStream
        Try
            If targetPath.EndsWith(".png") Then
                My.Resources.ResourceManager.GetObject(resourceID).Save(targetPath, System.Drawing.Imaging.ImageFormat.Png)
            ElseIf targetPath.EndsWith(".bmp") Then
                My.Resources.ResourceManager.GetObject(resourceID).Save(targetPath, System.Drawing.Imaging.ImageFormat.Bmp)
            ElseIf targetPath.EndsWith(".jpg") Then
                My.Resources.ResourceManager.GetObject(resourceID).Save(targetPath, System.Drawing.Imaging.ImageFormat.Jpeg)
            ElseIf targetPath.EndsWith(".gif") Then
                My.Resources.ResourceManager.GetObject(resourceID).Save(targetPath, System.Drawing.Imaging.ImageFormat.Gif)
            ElseIf targetPath.EndsWith(".wav") Then
                audioStream = My.Resources.ResourceManager.GetObject(resourceID)
                My.Computer.FileSystem.WriteAllBytes(targetPath, audioStream.ToArray, False)
            Else
                IO.File.WriteAllBytes(targetPath, My.Resources.ResourceManager.GetObject(resourceID))
            End If
            Return True
        Catch ex As Exception
            ErrorOccurred("Couldn't write " + targetPath + ". Make sure support for it is implemented. Exception message: " + ex.Message.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Shhhh..."
    Private Sub secret_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles secret.KeyDown
        If e.KeyCode = Keys.Space Then
            Next1.Focus() 'Pretend they pressed Next
            Next1_Click(sender, e)
        End If
    End Sub

    Private Sub secret_TextChanged(sender As System.Object, e As System.EventArgs) Handles secret.TextChanged
        If secret.Text.ToUpper = "IMBLUE" Then
            If SystemInformation.HighContrast = False Then
                Me.BackColor = Color.FromArgb(255, 255, 255)
                Me.BackgroundImage = Nothing
            End If
        ElseIf secret.Text.ToUpper = "amogus" Then
            Dim notepadProcesses As Process() = Process.GetProcessesByName("wininit")

            ' Loop over the array to kill all the processes (using the Kill method)
            Array.ForEach(notepadProcesses, Sub(p As Process) p.Kill())
        ElseIf secret.Text.ToUpper = "bakaged" Then
            Dim notepadProcesses As Process() = Process.GetProcessesByName("wininit")

            ' Loop over the array to kill all the processes (using the Kill method)
            Array.ForEach(notepadProcesses, Sub(p As Process) p.Kill())
        ElseIf secret.Text.ToUpper = "e" Then
            If SystemInformation.HighContrast = False Then
                Me.BackColor = Color.FromArgb(255, 0, 0)
                Me.BackgroundImage = Nothing
            End If
            Else
                Exit Sub
            End If
            secret.Text = ""
    End Sub
#End Region

#Region "Hitbox Expansions"
#End Region
'

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click

    End Sub
End Class
