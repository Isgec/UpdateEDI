Public Class frmMain
  Delegate Sub ThreadedSub()
  Delegate Sub ThreadedShow(ByVal slzFile As String)
  Delegate Sub ThreadedNone()
  Dim WithEvents jp As JobProcess.JobProcessor = Nothing
  Private Sub cmdStart_Click(sender As Object, e As EventArgs) Handles cmdStart.Click
    cmdStart.Enabled = False
    cmdStart.Text = "Loading..."
    ListBox1.Items.Clear()
    ListBox2.Items.Clear()
    Dim tmp As ThreadedSub = AddressOf Start
    tmp.BeginInvoke(Nothing, Nothing)
  End Sub
  Private Sub Start()
    jp = New JobProcess.JobProcessor()
    jp.Start()
  End Sub

  Private Sub cmdStop_Click(sender As Object, e As EventArgs) Handles cmdStop.Click
    cmdStop.Enabled = False
    cmdStop.Text = "Closing..."
    jp.StopJob()
  End Sub

  Private Sub jp_JobStarted() Handles jp.JobStarted
    If cmdStart.InvokeRequired Then
      cmdStart.Invoke(New ThreadedNone(AddressOf jp_JobStarted))
    Else
      cmdStart.Enabled = False
      cmdStart.Text = "Start"
      cmdStop.Enabled = True
    End If
  End Sub
  Private Sub jp_JobStopped() Handles jp.JobStopped
    If cmdStop.InvokeRequired Then
      cmdStop.Invoke(New ThreadedNone(AddressOf jp_JobStopped))
    Else
      cmdStop.Enabled = False
      cmdStop.Text = "Stop"
      cmdStart.Enabled = True
    End If
  End Sub

  Private Sub jp_Log(Message As String) Handles jp.Log
    If ListBox1.InvokeRequired Then
      ListBox1.Invoke(New ThreadedShow(AddressOf jp_Log), Message)
    Else
      ListBox1.Items.Insert(0, Message)
      Label1.Text = Message
    End If
  End Sub

  Private Sub jp_Err(Message As String) Handles jp.Err
    If ListBox2.InvokeRequired Then
      ListBox2.Invoke(New ThreadedShow(AddressOf jp_Err), Message)
    Else
      ListBox2.Items.Insert(0, Message)
      Label1.Text = Message
    End If
  End Sub
End Class