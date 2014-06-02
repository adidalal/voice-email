Imports Microsoft.Office.Tools.Ribbon
Imports System.IO
Imports System.Reflection

Public Class VoiceEmailAddin
    ' initialize objects
    Dim AudioRecorder As New Argus.Audio.Recording()
    Dim curCursor As Windows.Forms.Cursor
    Dim mailItem As Outlook.MailItem
    Dim strFile As String
    Dim isRecording As Boolean

    ' default button states
    Private Sub VoiceEmailAddin_Load(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs) Handles MyBase.Load
        btnStart.Enabled = True
        btnPause.Enabled = False
        btnStop.Enabled = False
        isRecording = False
    End Sub

    ' add form text 1 to message body
    Private Sub btnText_1_Click(sender As Object, e As RibbonControlEventArgs) Handles btnText_1.Click
        Dim strBody As String
        mailItem = TryCast(Globals.ThisAddIn.Application.ActiveInspector().CurrentItem, Outlook.MailItem)
        If Not (mailItem Is Nothing) Then
            If mailItem.EntryID Is Nothing Then
                ' add in message text
                strBody = mailItem.Body
                mailItem.Body = My.Settings("MessageText1") & vbCrLf & strBody
            End If
        End If
    End Sub

    ' add form text 2 to message body
    Private Sub btnText_2_Click(sender As Object, e As RibbonControlEventArgs) Handles btnText_2.Click
        Dim strBody As String
        mailItem = TryCast(Globals.ThisAddIn.Application.ActiveInspector().CurrentItem, Outlook.MailItem)
        If Not (mailItem Is Nothing) Then
            If mailItem.EntryID Is Nothing Then
                ' add in message text
                strBody = mailItem.Body
                mailItem.Body = My.Settings("MessageText2") & vbCrLf & strBody
            End If
        End If
    End Sub

    ' start Recording
    Private Sub btnStart_Click(sender As Object, e As RibbonControlEventArgs) Handles btnStart.Click
        ' update button states
        btnStart.Enabled = False
        btnStart.Label = "Recording"
        btnPause.Enabled = True
        btnStop.Enabled = True
        isRecording = True
        ' call AudioRecorder
        AudioRecorder.SamplesPerSecond = Argus.Audio.Recording.SamplesPerSecValue.LOW
        AudioRecorder.Filename = My.Settings("path") & "message.wav" ' temporary storage location for recorded message
        AudioRecorder.StartRecording()
    End Sub

    ' pause Recording
    Private Sub btnPause_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPause.Click
        If (isRecording) Then
            AudioRecorder.Pause = True
            ' update button states
            btnStart.Enabled = False
            btnStart.Label = "Paused"
            btnPause.Enabled = True
            btnPause.Label = "Resume"
            btnPause.Image = My.Resources.play
            btnStop.Enabled = False
            isRecording = False
        Else
            ' resume audio recording
            AudioRecorder.Pause = False
            ' update button states
            btnStart.Enabled = False
            btnStart.Label = "Recording..."
            btnPause.Enabled = True
            btnPause.Label = "Pause"
            btnPause.Image = My.Resources.pause
            btnStop.Enabled = True
            isRecording = True
        End If
    End Sub

    ' stop recording
    Private Sub btnStop_Click(sender As Object, e As RibbonControlEventArgs) Handles btnStop.Click
        ' change cursor to "busy"
        curCursor = Windows.Forms.Cursor.Current
        Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
        ' try stopping recording
        mailItem = TryCast(Globals.ThisAddIn.Application.ActiveInspector().CurrentItem, Outlook.MailItem)
        If Not (mailItem Is Nothing) Then
            ' update button states
            btnStart.Enabled = True
            btnStart.Label = "Start"
            btnPause.Enabled = False
            btnStop.Enabled = False
            isRecording = False
            ' stop recording
            AudioRecorder.StopRecording()
            ' convert to mp3
            Try
                Shell("cmd.exe /c " & My.Settings("path") & "lame.exe -a -b 32 " & My.Settings("path") & "message.wav " & My.Settings("path") & "message.mp3", AppWinStyle.Hide, True, -1)
                strFile = My.Settings("path") & "message.mp3"
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show("Conversion to MP3 Failed, attaching as WAV")
                strFile = My.Settings("path") & "message.wav"
            End Try
            ' attach file
            mailItem.Attachments.Add(strFile)
        End If

        ' reset cursor to default
        Windows.Forms.Cursor.Current = curCursor
    End Sub

    ' settings
    Private Sub btnSettings_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSettings.Click
        Dim frm As New userSettings
        frm.ShowDialog()

    End Sub

    ' help
    Private Sub btnHelp_Click(sender As Object, e As RibbonControlEventArgs) Handles btnHelp.Click
        System.Windows.Forms.MessageBox.Show("Help")
    End Sub

    ' about
    Private Sub btnAbout_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAbout.Click
        Dim frm As New frmAbout
        frm.ShowDialog()
    End Sub
End Class
