Public Class Form1
    Public WithEvents MIDI As New MidiFunctions
    Private Sub MIDI_MTCStatusChanged(sender As Object, e As MidiFunctions.MTCStatusChangedEventArgs) Handles MIDI.MTCStatusChanged
        Me.BeginInvoke(Sub()
                           Select Case e.Status
                               Case MidiFunctions.MTCStatus.Running
                                   Me.TextBox2.BackColor = Color.LightGreen

                               Case MidiFunctions.MTCStatus.Stopped
                                   Me.TextBox2.BackColor = Color.OrangeRed

                               Case MidiFunctions.MTCStatus.WaitingForFirstFrame
                                   Me.TextBox2.BackColor = Color.Yellow

                           End Select

                       End Sub)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.MIDI.MTCSenderStop()
        Me.MIDI.CloseMIDI()
        Me.MIDI.CloseMIDI_In()
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Load MIDI-OUT devices
        For Each device In MIDI.getDeviceList
            Me.ComboBox1.Items.Add(device)
        Next

        'Load MIDI-IN devices
        For Each device In MIDI.getDeviceList_In
            Me.ComboBox2.Items.Add(device)
        Next


        'Start UI Timer which refreshes the display values
        Me.UITimer.Start()
    End Sub

    Private WithEvents UITimer As New System.Timers.Timer(30)
    Private UITimerCounter1 As Int16 = 0
    Private Sub UITimer_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles UITimer.Elapsed
        'Updating UI display values
        Me.BeginInvoke(Sub()
                           Me.TextBox1.Text = MIDI.SendingMTC.ToStringWithSMPTE
                           Me.TextBox2.Text = MIDI.ReceivedMTC.ToStringWithSMPTE

                           If Me.MIDI.isMTCSenderRunning Then
                               Me.TextBox1.BackColor = Color.LightGreen
                           Else
                               Me.TextBox1.BackColor = Color.OrangeRed
                           End If
                       End Sub)
    End Sub


    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Me.MIDI.CloseMIDI()
        If Not String.IsNullOrWhiteSpace(Me.ComboBox1.SelectedItem.ToString) Then
            Me.MIDI.SetDevice(Me.ComboBox1.SelectedItem.ToString)
            Me.MIDI.OpenMIDI()
        End If
    End Sub
    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged
        Me.MIDI.CloseMIDI_In()
        If Not String.IsNullOrWhiteSpace(Me.ComboBox2.SelectedItem.ToString) Then
            Me.MIDI.SetDevice_In(Me.ComboBox2.SelectedItem.ToString)
            Me.MIDI.OpenMIDI_In()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.MIDI.isMTCSenderRunning Then
            'Restarting MTC Sender
            Dim mtc = New MidiFunctions.MTCFrame With {.FrameRate = MidiFunctions.MTCFrameRate.fps30}
            Me.MIDI.MTCSenderGoto(mtc)

        Else
            'Starting MTC Sender
            Me.MIDI.MTCSenderStart()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.MIDI.MTCSenderStop()
        Me.MIDI.MTCSenderGoto(New MidiFunctions.MTCFrame With {.FrameRate = MidiFunctions.MTCFrameRate.fps30})
    End Sub
End Class