Public Class PrepForm
#Region "Classic theme, because Classic theme's funny"
    <DllImport("uxtheme.dll", CharSet:=CharSet.Auto)> _
    Public Shared Sub SetThemeAppProperties(ByVal Flags As Integer)
    End Sub
#End Region

#Region "Variables"
    Private ex7forw8notice1 As String = <a>This option is only available on
Windows Vista.0.</a>
    Private ex7forw8notice2 As String = <a>This option breaks if used with
high contrast.</a>
#End Region

    Private Sub PrepForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'First, error out on incompatible versions
        If Not System.Environment.OSVersion.Version.Major = 6 Or (System.Environment.OSVersion.Version.Minor > 0) Then
            Me.Hide()
            SetThemeAppProperties(0)
            MsgBox("This transformation pack is only compatible with Windows Vista Extended Kernel.", MsgBoxStyle.Critical, "Unsupported OS")
            End
        End If
        If Environment.Is64BitOperatingSystem = False Then
            MsgBox("This transformation pack is only compatible with Windows Vista Extended Kernel.", MsgBoxStyle.Critical, "Unsupported OS")
            End
        End If

        Me.Show()
        Me.Refresh()

        Form1.RefreshHC() 'Manually trigger HC check, since it otherwise ONLY gets triggered by the theme changing

        If Not Form1.HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer") Is Nothing Then 'We're in a transformed machine?
            Try
                'Remaining prefs
                Form1.AllowUXThemePatcher.Checked = (Form1.HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer", True).GetValue("AllowUXThemePatcher") = "true")
                Form1.AllowUXThemePatcher.Enabled = False 'prevent changing its... install-state to prevent a brick

                If Form1.HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("CurrentPhase") = 10 Then 'Load configuration
                    Form1.PageOptions.Visible = True
                    Form1.Page1.Visible = False
                    Form1.BackArea.Visible = True
                    Form1.isInstalled = True
                    Form1.NextConfirm.Text = "Apply"
                    Form1.Text = "Manage - Longhornizer Transformation Pack "
                    Form1.Label17.Text = "Select extra components"
                    Form1.Label16.Text = "Manage the extra components installed on your computer below. Once done, click Apply to apply your changes."
                    Form1.Label12.Visible = False 'Hide the transformation notice
                    Form1.Label20.Text = "Updating transformation..."
                End If
            Catch
            End Try
        End If

        ProgressBar1.Visible = True
        Dim jobthread As New Thread(AddressOf LoadPack)
        jobthread.Start()
    End Sub

    Private Sub LoadPack()
        If Form1.HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer") Is Nothing Then
            ChangeStatus("Please wait...")
            ' Warning message for non-extended kernel users
            Try
                If MsgBox("This Transformation Pack is meant for Windows Vista with the extended kernel installed. Your system will become unbootable if you don't have it. Go to ximonite.com/win32 to download it" + Environment.NewLine + Environment.NewLine + "KNOWING THIS, would you like to continue?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.SystemModal, "Warning") = MsgBoxResult.Yes Then
                    SetThemeAppProperties(3)
                Else
                    End
                End If
            Catch ex As Exception
                If MsgBox("This Transformation Pack is meant for Windows Vista with the extended kernel installed. Your system will become unbootable if you don't have it. Go to ximonite.com/win32 to download it" + Environment.NewLine + Environment.NewLine + "KNOWING THIS, would you like to continue?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.SystemModal, "Warning") = MsgBoxResult.Yes Then
                    SetThemeAppProperties(3)
                End If
            End Try
        End If
        ChangeStatus("Loading...")
        CloseSplash()
    End Sub

    Private Sub CloseSplash()
        'This bit here makes it work despite being called via a VB.NET Thread
        If Me.InvokeRequired Then
            Me.Invoke(New Action(AddressOf CloseSplash))
            Exit Sub
        End If

        Form1.Show()
        Form1.secret.Focus()
        Me.Close()
    End Sub

    Private Sub ChangeStatus(ByVal status As String)
        'This bit here makes it work despite being called via a VB.NET Thread
        If Me.InvokeRequired Then
            Dim args() As String = {status}
            Me.Invoke(New Action(Of String)(AddressOf ChangeStatus), args)
            Exit Sub
        End If

        Me.Text = status
    End Sub
End Class
