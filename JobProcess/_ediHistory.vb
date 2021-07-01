Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.EDI
  <DataObject()> _
  Partial Public Class ediHistory
    Private Shared _RecordCount As Integer
    Public Property HistoryNo As Int32 = 0
    Public Property SerialNo As Int32 = 0
    Public Property EdiKey As String = ""
    Public Property EdiValues As String = ""
    Public Property ExecutedStatement As String = ""
    Private _ExecutedOn As String = ""
    Public Property ExecutedOn() As String
      Get
        If Not _ExecutedOn = String.Empty Then
          Return Convert.ToDateTime(_ExecutedOn).ToString("dd/MM/yyyy HH:mm:ss")
        End If
        Return _ExecutedOn
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _ExecutedOn = ""
        Else
          _ExecutedOn = value
        End If
      End Set
    End Property
    Public Readonly Property DisplayField() As String
      Get
        Return ""
      End Get
    End Property
    Public Readonly Property PrimaryKey() As String
      Get
        Return _HistoryNo
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
    Public Class PKediHistory
      Private _HistoryNo As Int32 = 0
      Public Property HistoryNo() As Int32
        Get
          Return _HistoryNo
        End Get
        Set(ByVal value As Int32)
          _HistoryNo = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function ediHistoryGetNewRecord() As SIS.EDI.ediHistory
      Return New SIS.EDI.ediHistory()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function ediHistoryGetByID(ByVal HistoryNo As Int32) As SIS.EDI.ediHistory
      Dim Results As SIS.EDI.ediHistory = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediHistorySelectByID"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@HistoryNo",SqlDbType.Int,HistoryNo.ToString.Length, HistoryNo)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.EDI.ediHistory(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function ediHistorySelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.EDI.ediHistory)
      Dim Results As List(Of SIS.EDI.ediHistory) = Nothing
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          If OrderBy = String.Empty Then OrderBy = "HistoryNo DESC"
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spediHistorySelectListSearch"
            SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spediHistorySelectListFilteres"
          End If
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.EDI.ediHistory)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.EDI.ediHistory(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function ediHistorySelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
      Return _RecordCount
    End Function
      'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Insert, True)> _
    Public Shared Function ediHistoryInsert(ByVal Record As SIS.EDI.ediHistory) As SIS.EDI.ediHistory
      Dim _Rec As SIS.EDI.ediHistory = SIS.EDI.ediHistory.ediHistoryGetNewRecord()
      With _Rec
        .SerialNo = Record.SerialNo
        .EdiKey = Record.EdiKey
        .EdiValues = Record.EdiValues
        .ExecutedStatement = Record.ExecutedStatement
        .ExecutedOn = Record.ExecutedOn
      End With
      Return SIS.EDI.ediHistory.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.EDI.ediHistory) As SIS.EDI.ediHistory
      Using Con As SqlConnection = New SqlConnection(SIS.SYS.SQLDatabase.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spediHistoryInsert"
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@SerialNo",SqlDbType.Int,11, Record.SerialNo)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@EdiKey",SqlDbType.NVarChar,51, Record.EdiKey)
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@EdiValues",SqlDbType.NVarChar,2147484, Iif(Record.EdiValues= "" ,Convert.DBNull, Record.EdiValues))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ExecutedStatement",SqlDbType.NVarChar,2147484, Iif(Record.ExecutedStatement= "" ,Convert.DBNull, Record.ExecutedStatement))
          SIS.SYS.SQLDatabase.DBCommon.AddDBParameter(Cmd, "@ExecutedOn",SqlDbType.DateTime,21, Iif(Record.ExecutedOn= "" ,Convert.DBNull, Record.ExecutedOn))
          Cmd.Parameters.Add("@Return_HistoryNo", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_HistoryNo").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.HistoryNo = Cmd.Parameters("@Return_HistoryNo").Value
        End Using
      End Using
      Return Record
    End Function
    Public Sub New(ByVal Reader As SqlDataReader)
      SIS.SYS.SQLDatabase.DBCommon.NewObj(Me, Reader)
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace
