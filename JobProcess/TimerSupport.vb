Imports System.Timers
Public MustInherit Class TimerSupport
  Private _tmr As Timer
  Private _stop As Boolean = False
  Public Property Interval() As Integer = 1000
  Public ReadOnly Property IsStopping() As Boolean
    Get
      Return _stop
    End Get
  End Property
  Public Sub Start()
    _tmr = New Timer
    AddHandler _tmr.Elapsed, New ElapsedEventHandler(AddressOf TimerFirst)
    _tmr.Interval = _Interval
    _tmr.Enabled = True
  End Sub
  Public Sub StopJob()
    _stop = True
  End Sub
  Private Sub TimerFirst(ByVal source As Object, ByVal e As ElapsedEventArgs)
    _tmr.Enabled = False
    RemoveHandler _tmr.Elapsed, New ElapsedEventHandler(AddressOf TimerFirst)
    If _stop Then
      _tmr = Nothing
      Try
        Stopped()
      Catch ex As Exception
      End Try
      Exit Sub
    End If
    Try
      Started()
    Catch ex As Exception
    End Try
    Try
      Process()
    Catch ex As Exception
    End Try
    AddHandler _tmr.Elapsed, New ElapsedEventHandler(AddressOf TimerTick)
    _tmr.Enabled = True
  End Sub
  Private Sub TimerTick(ByVal source As Object, ByVal e As ElapsedEventArgs)
    _tmr.Enabled = False
    If _stop Then
      _tmr = Nothing
      Try
        Stopped()
      Catch ex As Exception
      End Try
      Exit Sub
    End If
    Try
      Process()
    Catch ex As Exception
    End Try
    _tmr.Enabled = True
  End Sub
  Public MustOverride Sub Process()
  Public MustOverride Sub Started()
  Public MustOverride Sub Stopped()
End Class
