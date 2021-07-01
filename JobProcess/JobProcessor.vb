Imports System.Reflection
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Data
Imports System.Data.SqlClient

Public Class JobProcessor
  Inherits TimerSupport
  Implements IDisposable

  Private jpConfig As ConfigFile = Nothing
  Private MayContinue As Boolean = True
  Private LibraryPath As String = ""
  Private LibraryID As String = ""
  Private RemoteLibraryConnected As Boolean = False

  Public Event JobStarted()
  Public Event JobStopped()
  Public Event Log(ByVal Message As String)
  Public Event Err(ByVal Message As String)

  Public Sub Msg(ByVal str As String)
    RaiseEvent Log(str)
  End Sub

  Public Sub MsgErr(ByVal str As String)
    RaiseEvent Err(str)
  End Sub
  Public Overrides Sub Process()
    Try
      Dim Jobs As List(Of SIS.EDI.ediQueues) = SIS.EDI.ediQueues.ediQueuesSelectList(0, 10, "", False, "", "")
      If Jobs.Count > 0 Then
        Msg(Jobs.Count & " record found for processing.")
      End If
      For Each job As SIS.EDI.ediQueues In Jobs
        Msg("Processing: " & job.SerialNo & " " & job.EdiKey)
        Try
          ProcessJob(job)
          Msg("Processed: " & job.SerialNo & " " & job.EdiKey)
        Catch ex As Exception
          MsgErr(ex.Message)
        End Try
        If IsStopping Then
          Exit For
        End If
      Next
      If IsStopping Then
        Msg("Cancelled")
      End If
    Catch ex As Exception
      MsgErr(ex.Message)
    End Try
  End Sub
  Public Sub ProcessJob(Job As SIS.EDI.ediQueues)
    Dim Key As SIS.EDI.ediKeys = Job.FK_EDI_Queues_EDIKey
    Dim para As List(Of String) = Key.EdiParameters.ToLower.Split("|".ToCharArray).ToList
    Dim value As List(Of String) = Job.EdiValues.Split("|".ToCharArray).ToList
    Dim sql As String = ""
    If para.Count <> value.Count Then
      Throw New Exception("Parameter Vs Values count mismatch.")
    End If
    sql = Key.SqlStatement.ToLower
    If sql.IndexOf("delete") >= 0 Then
      Throw New Exception("DELETE SQL is NOT Allowed.")
    End If
    If Key.IsSP Then
    Else
      '1. Create SQL Statement By Values
      For I As Integer = 0 To para.Count - 1
        Dim s As String = para(I)
        Dim v As String = value(I)
        sql = sql.Replace(s, v)
      Next
      '2. Standard Replacements in SQL
      sql = sql.Replace("@comp", Key.ERPCompany)
      '3. Create History Record before execution
      Dim hist As New SIS.EDI.ediHistory
      With hist
        .EdiKey = Job.EdiKey
        .EdiValues = Job.EdiValues
        .ExecutedOn = Now.ToString("dd/MM/yyyy HH:mm:ss")
        .ExecutedStatement = sql
        .SerialNo = Job.SerialNo
      End With
      SIS.EDI.ediHistory.InsertData(hist)
      '4. Execute SQL
      If Key.ExecuteInERP Then
        Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = sql
            Con.Open()
            Cmd.ExecuteNonQuery()
          End Using
        End Using
      Else
        Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString(Key.ERPCompany))
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = sql
            Con.Open()
            Cmd.ExecuteNonQuery()
          End Using
        End Using
      End If
      '5. Delete job, if successful
      SIS.EDI.ediQueues.ediQueuesDelete(Job)
    End If
  End Sub

  Public Overrides Sub Started()
    Try
      RaiseEvent JobStarted()
      Msg("Reading Settings")
      Dim ConfigPath As String = IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\Settings.xml"
      jpConfig = ConfigFile.DeSerialize(ConfigPath)
      SIS.SYS.SQLDatabase.DBCommon.BaaNLive = jpConfig.BaaNLive
      SIS.SYS.SQLDatabase.DBCommon.JoomlaLive = jpConfig.JoomlaLive
    Catch ex As Exception
      StopJob()
      MsgErr(ex.Message)
    End Try
  End Sub

  Public Overrides Sub Stopped()
    jpConfig = Nothing
    RaiseEvent JobStopped()
    Msg("Stopped")
  End Sub

  Public Shared Function IsFileAvailable(ByVal FilePath As String) As Boolean
    If Not IO.File.Exists(FilePath) Then Return False
    Dim fInfo As IO.FileInfo = Nothing
    Dim st As IO.FileStream = Nothing
    Try
      fInfo = New IO.FileInfo(FilePath)
    Catch ex As Exception
      Return False
    End Try
    Dim ret As Boolean = False
    If fInfo.IsReadOnly Then
      If DateDiff(DateInterval.Minute, fInfo.CreationTime, Now) >= 1 Then
        fInfo.IsReadOnly = False
      End If
    End If
    Try
      st = fInfo.Open(IO.FileMode.Open, IO.FileAccess.ReadWrite, IO.FileShare.None)
      ret = True
    Catch ex As Exception
      ret = False
    Finally
      If st IsNot Nothing Then
        st.Close()
      End If
    End Try
    Return ret
  End Function

  Sub New()
    'dummy
  End Sub
  ''' <summary>
  ''' Enter connection settings for BaaN and Joomla
  ''' </summary>
  ''' <param name="BaaNLive"></param>
  ''' <param name="JoomlaLive"></param>
  Sub New(BaaNLive As Boolean, JoomlaLive As Boolean)
    SIS.SYS.SQLDatabase.DBCommon.BaaNLive = BaaNLive
    SIS.SYS.SQLDatabase.DBCommon.JoomlaLive = JoomlaLive
  End Sub

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub
#End Region
End Class
