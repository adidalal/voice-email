Partial Class VoiceEmailAddin
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.VoiceEmail = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btnText_1 = Me.Factory.CreateRibbonButton
        Me.btnText_2 = Me.Factory.CreateRibbonButton
        Me.btnStart = Me.Factory.CreateRibbonButton
        Me.btnPause = Me.Factory.CreateRibbonButton
        Me.btnStop = Me.Factory.CreateRibbonButton
        Me.btnSettings = Me.Factory.CreateRibbonButton
        Me.btnHelp = Me.Factory.CreateRibbonButton
        Me.btnAbout = Me.Factory.CreateRibbonButton
        Me.VoiceEmail.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        '
        'VoiceEmail
        '
        Me.VoiceEmail.Groups.Add(Me.Group1)
        Me.VoiceEmail.Groups.Add(Me.Group2)
        Me.VoiceEmail.Groups.Add(Me.Group3)
        Me.VoiceEmail.Label = "Voice Email"
        Me.VoiceEmail.Name = "VoiceEmail"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btnText_1)
        Me.Group1.Items.Add(Me.Separator1)
        Me.Group1.Items.Add(Me.btnText_2)
        Me.Group1.Label = "Insert Text"
        Me.Group1.Name = "Group1"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnStart)
        Me.Group2.Items.Add(Me.btnPause)
        Me.Group2.Items.Add(Me.btnStop)
        Me.Group2.Label = "Record"
        Me.Group2.Name = "Group2"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnSettings)
        Me.Group3.Items.Add(Me.btnHelp)
        Me.Group3.Items.Add(Me.btnAbout)
        Me.Group3.Label = "Configuration"
        Me.Group3.Name = "Group3"
        '
        'btnText_1
        '
        Me.btnText_1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnText_1.Image = Global.VoiceRecorder_Outlook.My.Resources.Resources.text1
        Me.btnText_1.Label = "Text 1"
        Me.btnText_1.Name = "btnText_1"
        Me.btnText_1.ShowImage = True
        '
        'btnText_2
        '
        Me.btnText_2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnText_2.Image = Global.VoiceRecorder_Outlook.My.Resources.Resources.text2
        Me.btnText_2.Label = "Text 2"
        Me.btnText_2.Name = "btnText_2"
        Me.btnText_2.ShowImage = True
        '
        'btnStart
        '
        Me.btnStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnStart.Image = Global.VoiceRecorder_Outlook.My.Resources.Resources.start
        Me.btnStart.Label = "Start"
        Me.btnStart.Name = "btnStart"
        Me.btnStart.ShowImage = True
        '
        'btnPause
        '
        Me.btnPause.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnPause.Enabled = False
        Me.btnPause.Image = Global.VoiceRecorder_Outlook.My.Resources.Resources.pause
        Me.btnPause.Label = "Pause"
        Me.btnPause.Name = "btnPause"
        Me.btnPause.ShowImage = True
        '
        'btnStop
        '
        Me.btnStop.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnStop.Enabled = False
        Me.btnStop.Image = Global.VoiceRecorder_Outlook.My.Resources.Resources._stop
        Me.btnStop.Label = "Stop"
        Me.btnStop.Name = "btnStop"
        Me.btnStop.ShowImage = True
        '
        'btnSettings
        '
        Me.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSettings.Image = Global.VoiceRecorder_Outlook.My.Resources.Resources.settings
        Me.btnSettings.Label = "Settings"
        Me.btnSettings.Name = "btnSettings"
        Me.btnSettings.ShowImage = True
        '
        'btnHelp
        '
        Me.btnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnHelp.Image = Global.VoiceRecorder_Outlook.My.Resources.Resources.help
        Me.btnHelp.Label = "Help"
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.ShowImage = True
        Me.btnHelp.Visible = False
        '
        'btnAbout
        '
        Me.btnAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAbout.Image = Global.VoiceRecorder_Outlook.My.Resources.Resources.about
        Me.btnAbout.Label = "About"
        Me.btnAbout.Name = "btnAbout"
        Me.btnAbout.ShowImage = True
        '
        'VoiceEmailAddin
        '
        Me.Name = "VoiceEmailAddin"
        Me.RibbonType = "Microsoft.Outlook.Mail.Compose"
        Me.Tabs.Add(Me.VoiceEmail)
        Me.VoiceEmail.ResumeLayout(False)
        Me.VoiceEmail.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()

    End Sub

    Friend WithEvents VoiceEmail As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents btnText_1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnStart As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnStop As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAbout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnPause As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnText_2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents btnSettings As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnHelp As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As VoiceEmailAddin
        Get
            Return Me.GetRibbon(Of VoiceEmailAddin)()
        End Get
    End Property
End Class
