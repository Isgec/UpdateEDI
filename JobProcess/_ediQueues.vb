Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.EDI
  <DataObject()> _
  Partial Public Class ediQueues
    Private Shared _RecordCount As Integer
    Public Property SerialNo As Int32 = 0
    Public Property EdiKey As String = ""
    Public Property EdiValues As String = ""
    Public Property EDI_Keys1_EdiParameters As String = ""
    Private _FK_EDI_Queues_EDIKey As SIS.EDI.ediKeys = Nothing
    Public ReadOnly Property DisplayField() As String
      Get
        Return ""
      End Get
    End Property
    Public ReadOnly Property PrimaryKey() As String
      Get
        Return _SerialNo
      End Get
    End Property
    Public Shared Property RecordCount() As Integer
      Get
        Return _RecordCount
      End Get
      Set(ByVal value As Integer)
        _RecordCount = value
      End Set
    End Property
    Public Class PKediQueues
      Private _SerialNo As Int32 = 0
      Public Property SerialNo() As Int32
        Get
          Return _SerialNo
        End Get
        Set(ByVal value As Int32)
          _SerialNo = value
        End Set
      End Property
    End Class
    Public ReadOnly Property FK_EDI_Queues_EDIKey() As SIS.EDI.ediKeys
      Get
        If _FK_EDI_Queues_EDIKey Is Nothing Then
          _FK_EDI_Queues_EDIKey = SIS.EDI.ediKeys.ediKeysGetByID(_EdiKey)
        End If
        Return _FK_EDI_Queues_EDIKey
      End Get
    End Property
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediQueuesGetNewRecord() As SIS.EDI.ediQueues
      Return New SIS.EDI.ediQueues()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediQueuesGetByID(ByVal SerialNo As Int32) As SIS.EDI.ediQueues
      Dim Results As SIS.EDI.ediQueues = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediQueuesSelectByID"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SerialNo", SqlDbType.Int, SerialNo.ToString.Length, SerialNo)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.EDI.ediQueues(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function ediQueuesSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String, ByVal EdiKey As String) As List(Of SIS.EDI.ediQueues)
      Dim Results As List(Of SIS.EDI.ediQueues) = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          If OrderBy = String.Empty Then OrderBy = "SerialNo DESC"
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spediQueuesSelectListSearch"
            SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spediQueuesSelectListFilteres"
            SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Filter_EdiKey", SqlDbType.NVarChar, 50, IIf(EdiKey Is Nothing, String.Empty, EdiKey))
          End If
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.EDI.ediQueues)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.EDI.ediQueues(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function ediQueuesSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String, ByVal EdiKey As String) As Integer
      Return _RecordCount
    End Function
      'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function ediQueuesGetByID(ByVal SerialNo As Int32, ByVal Filter_EdiKey As String) As SIS.EDI.ediQueues
      Return ediQueuesGetByID(SerialNo)
    End Function
    <DataObjectMethod(DataObjectMethodType.Insert, True)> _
    Public Shared Function ediQueuesInsert(ByVal Record As SIS.EDI.ediQueues) As SIS.EDI.ediQueues
      Dim _Rec As SIS.EDI.ediQueues = SIS.EDI.ediQueues.ediQueuesGetNewRecord()
      With _Rec
        .EdiKey = Record.EdiKey
        .EdiValues = Record.EdiValues
      End With
      Return SIS.EDI.ediQueues.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.EDI.ediQueues) As SIS.EDI.ediQueues
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediQueuesInsert"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@EdiKey",SqlDbType.NVarChar,51, Record.EdiKey)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@EdiValues",SqlDbType.NVarChar,2147484, Iif(Record.EdiValues= "" ,Convert.DBNull, Record.EdiValues))
          Cmd.Parameters.Add("@Return_SerialNo", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_SerialNo").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.SerialNo = Cmd.Parameters("@Return_SerialNo").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)> _
    Public Shared Function ediQueuesUpdate(ByVal Record As SIS.EDI.ediQueues) As SIS.EDI.ediQueues
      Dim _Rec As SIS.EDI.ediQueues = SIS.EDI.ediQueues.ediQueuesGetByID(Record.SerialNo)
      With _Rec
        .EdiKey = Record.EdiKey
        .EdiValues = Record.EdiValues
      End With
      Return SIS.EDI.ediQueues.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.EDI.ediQueues) As SIS.EDI.ediQueues
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediQueuesUpdate"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Original_SerialNo",SqlDbType.Int,11, Record.SerialNo)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@EdiKey",SqlDbType.NVarChar,51, Record.EdiKey)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@EdiValues",SqlDbType.NVarChar,2147484, Iif(Record.EdiValues= "" ,Convert.DBNull, Record.EdiValues))
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Con.Open()
          Cmd.ExecuteNonQuery()
          _RecordCount = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Delete, True)> _
    Public Shared Function ediQueuesDelete(ByVal Record As SIS.EDI.ediQueues) As Int32
      Dim _Result as Integer = 0
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediQueuesDelete"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@Original_SerialNo",SqlDbType.Int,Record.SerialNo.ToString.Length, Record.SerialNo)
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Con.Open()
          Cmd.ExecuteNonQuery()
          _RecordCount = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return _RecordCount
    End Function
    Public Sub New(ByVal Reader As SqlDataReader)
      SIS.SYS.SQLDatabase.DBCommon.NewObj(Me, Reader)
    End Sub
    Public Sub New()
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
  End Class
End Namespace
