Imports Utilities.Utilities
Imports AdoDALInterface
Imports AdoDALInterface.AdoConstants
Imports System.Data.SqlClient

Public Class AdoDALSQL
    Implements IAdoDAL


    Public Property DatabaseInfo As AdoDALInterface.IDataBaseInfo Implements AdoDALInterface.IAdoDAL.DatabaseInfo
    Public Property DataBaseErrorMessages As List(Of String) = New List(Of String) Implements IAdoDAL.DataBaseErrorMessages
    Public Property DataBaseMessages As List(Of String) = New List(Of String) Implements IAdoDAL.DataBaseMessages
    Public Property oAdoDAL As IAdoDAL

    'Private CurrentDAL As AdoDALSQL = Me

    Private WithEvents oSQLConnection As System.Data.SqlClient.SqlConnection = Nothing
    Private oSQLCommand As System.Data.SqlClient.SqlCommand = Nothing
    Private oSQLDataAdapter As System.Data.SqlClient.SqlDataAdapter = Nothing
    Private oSQLTransaction As System.Data.SqlClient.SqlTransaction = Nothing
    Private bInTransaction As Boolean = False

    Public Overridable Sub PartialData(ByVal firstTime As Boolean, ResultsData As DataTable)

    End Sub

    Private ReadOnly Property SQLConnection() As System.Data.SqlClient.SqlConnection
        Get
            Try
                If Equal(oSQLCommand, Nothing) OrElse Equal(oSQLConnection, Nothing) OrElse oSQLConnection.State = ConnectionState.Broken OrElse oSQLConnection.State = ConnectionState.Closed Then
                    If Equal(oSQLTransaction, Nothing) Then Connect() Else DataBaseErrorMessages.Add("SQLConnection is lost while saving. Please try saving again")
                End If
            Catch ex As Exception
                Throw ex
            End Try
            Return oSQLConnection
        End Get
    End Property
    Private ReadOnly Property SQLCommand() As System.Data.SqlClient.SqlCommand
        Get
            Try
                If Equal(oSQLCommand, Nothing) OrElse Equal(oSQLConnection, Nothing) OrElse oSQLConnection.State = ConnectionState.Broken OrElse oSQLConnection.State = ConnectionState.Closed Then
                    If Equal(oSQLTransaction, Nothing) Then Connect() Else DataBaseErrorMessages.Add("SQLConnection is lost while saving. Please try saving again")
                End If
            Catch ex As Exception
                Throw ex
            End Try
            Return oSQLCommand
        End Get
    End Property
    Private ReadOnly Property SQLDataAdapter() As System.Data.SqlClient.SqlDataAdapter
        Get
            Try
                If Equal(oSQLCommand, Nothing) OrElse Equal(oSQLConnection, Nothing) OrElse oSQLConnection.State = ConnectionState.Broken OrElse oSQLConnection.State = ConnectionState.Closed Then
                    If Equal(oSQLTransaction, Nothing) Then Connect() Else DataBaseErrorMessages.Add("SQLConnection is lost while saving. Please try saving again")
                End If
            Catch ex As Exception
                Throw ex
            End Try
            Return oSQLDataAdapter
        End Get
    End Property
    Private ReadOnly Property SQLTransaction() As System.Data.SqlClient.SqlTransaction
        Get
            Try
                If Equal(oSQLCommand, Nothing) OrElse Equal(oSQLConnection, Nothing) OrElse oSQLConnection.State = ConnectionState.Broken OrElse oSQLConnection.State = ConnectionState.Closed Then
                    If Equal(oSQLTransaction, Nothing) Then Connect() Else DataBaseErrorMessages.Add("SQLConnection is lost while saving. Please try saving again")
                End If
            Catch ex As Exception
                Throw ex
            End Try
            Return oSQLTransaction
        End Get
    End Property
    Public Function Connect() As System.Data.Common.DbConnection Implements IAdoDAL.Connect
        Return Connect("Apex")
    End Function
    Sub Con_InfoMessage(sender As Object, e As SqlInfoMessageEventArgs) Handles oSQLConnection.InfoMessage
        Dim myMsg As String = e.Message
        DataBaseMessages.Add(myMsg)
    End Sub
    Public Function Connect(sApplicationName As String) As IDbConnection Implements IAdoDAL.Connect
        Dim bSuccess As Boolean = False
        Try
            Dim sPassword As String = String.Empty

            If Not Equal(DatabaseInfo.DatabasePassword, Nothing) Then
                If DatabaseInfo.DatabasePassword.IndexOf("~") > 0 Then ''Changed for Web
                    sPassword = Decrypt(DatabaseInfo.DatabasePassword)
                Else
                    sPassword = DatabaseInfo.DatabasePassword
                End If
            End If
            Dim sConnectionString As String = "Application Name=" & sApplicationName & ";Data Source=" & DatabaseInfo.DatabaseServerName & ";initial catalog=" & DatabaseInfo.DatabaseName & ";persist security info=True;packet size=4096;uid=" & DatabaseInfo.DatabaseUserName & ";pwd=" & sPassword
            oSQLConnection = New System.Data.SqlClient.SqlConnection(sConnectionString)
            AddHandler oSQLConnection.InfoMessage, AddressOf Con_InfoMessage

            Dim sCommandText As String = Nothing
            sCommandText = "SET DATEFORMAT 'DMY'"
            Try
                oSQLConnection.Open()
            Catch ex As Exception
                ''Added by G.K (v 4.6.0) If application is not there then it should log or throw error
                'If Not Equal(Base.BaseApplication.oApplication, Nothing) Then
                DataBaseErrorMessages.Add("Unable to establish database connection." & vbCrLf & "Please check database configuration.")
                ''Base.BaseApplication.oApplication.ExitApplication()
                'Exit Function
                'Else
                'Throw ex
                'End If
            End Try
            If oSQLConnection.State = ConnectionState.Open Then
                oSQLCommand = New System.Data.SqlClient.SqlCommand(sCommandText, oSQLConnection)
                oSQLCommand.CommandTimeout = 0
                oSQLDataAdapter = New System.Data.SqlClient.SqlDataAdapter(oSQLCommand)
                oSQLCommand.ExecuteNonQuery()
                DatabaseInfo.ConnectionID = Me.ExecuteScalar("SELECT @@SPID PID")
            End If
            DatabaseInfo.Connection = oSQLConnection
            bSuccess = True
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        Return oSQLConnection
        'Return Connect()
    End Function
    Public Function Close() As Boolean Implements IAdoDAL.Close
        Dim bSuccess As Boolean = False
        Try
            If oSQLConnection.State = ConnectionState.Open Then oSQLConnection.Close()
            bSuccess = True
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        Return bSuccess
    End Function

    Public Overridable Function ExecuteQuery(ByVal sQuery As String) As DataTable Implements IAdoDAL.ExecuteQuery
        Dim dtDataTable As DataTable = New DataTable
        Try
            SQLCommand.CommandText = sQuery
            SQLCommand.CommandType = CommandType.Text
            SQLCommand.Transaction = SQLTransaction
            SQLCommand.Connection = SQLConnection

            SQLDataAdapter.Fill(dtDataTable)
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        Return dtDataTable
    End Function

    Public Overridable Sub ExecuteQuery(ByVal sQuery As String, ByVal iRecordLimit As Integer) Implements IAdoDAL.ExecuteQuery 'As DataTable 
        Dim dtDataTable As DataTable = New DataTable
        Dim oSQLReader As SqlDataReader = SQLCommand.ExecuteReader()
        Try
            ''Dim iRecordsAffected As Integer = Me.ExecuteScalar("SELECT COUNT(*) FROM (" & sQuery & ") AS Q")

            'If iRecordsAffected < iRecordLimit Then
            '    dtDataTable = Me.ExecuteQuery(sQuery)
            'Else
            SQLCommand.CommandText = sQuery
            SQLCommand.CommandType = CommandType.Text
            SQLCommand.Transaction = SQLTransaction
            SQLCommand.Connection = SQLConnection


            Dim ListOfColumn As List(Of DataColumn) = BuildDataTable(oSQLReader, dtDataTable)
            ''CallReader 
            ReadSQLDataReader(oSQLReader, dtDataTable, iRecordLimit, ListOfColumn)
            oSQLReader.Close()
            ''End If
        Catch ex As Exception
            oSQLReader.Close()
            LogSystemError(ex)
            Throw ex
        End Try
        ''Return dtDataTable
    End Sub

    Public Overridable Function ExecuteNonQueryAsync(sQuery As String) As Integer Implements IAdoDAL.ExecuteNonQueryAsync
        SQLCommand.CommandText = sQuery
        SQLCommand.CommandType = CommandType.Text
        SQLCommand.Transaction = SQLTransaction
        SQLCommand.Connection = SQLConnection
        Dim result = SQLCommand.BeginExecuteNonQuery(Sub(p)
                                                         Try
                                                             Dim asyncCommand = TryCast(p.AsyncState, SqlCommand)
                                                             Console.WriteLine("Execution Completed")
                                                         Catch ex As Exception
                                                             Console.WriteLine("Error:::{0}", ex.Message)
                                                         Finally
                                                             ''conn.Close()

                                                         End Try
                                                     End Sub, SQLCommand)
        Return SQLCommand.EndExecuteNonQuery(result)
    End Function

    Public Function ExecuteNonQuery(ByVal sQuery As String) As Integer Implements IAdoDAL.ExecuteNonQuery
        Dim iRowsAffected As Integer = 0
        Try
            SQLCommand.CommandText = sQuery
            SQLCommand.CommandType = CommandType.Text
            SQLCommand.Transaction = SQLTransaction
            SQLCommand.Connection = SQLConnection
            iRowsAffected = SQLCommand.ExecuteNonQuery
        Catch ex As Exception
            'LogSystemError(ex)
            Throw ex
        End Try
        Return iRowsAffected
    End Function
    Public Function ExecuteScalar(ByVal sQuery As String) As Object Implements IAdoDAL.ExecuteScalar
        Dim oObject As Object = Nothing
        Try
            SQLCommand.CommandText = sQuery
            SQLCommand.CommandType = CommandType.Text
            SQLCommand.Transaction = SQLTransaction
            SQLCommand.Connection = SQLConnection
            oObject = SQLCommand.ExecuteScalar
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        If Equal(oObject, Nothing) Then
            Return Nothing
        Else
            Return oObject
        End If
    End Function
    Public Function ExecuteProcedure(ByVal sProcedureName As String, ByVal Parameters()() As Object) As DataTable Implements IAdoDAL.ExecuteProcedure
        'Columns of array are: sParameterName, sDataType, iSize, oValue
        Dim dtDataTable As DataTable = New DataTable
        Try
            SQLCommand.CommandText = sProcedureName
            SQLCommand.CommandType = CommandType.StoredProcedure
            SQLCommand.Transaction = SQLTransaction
            SQLCommand.Connection = SQLConnection
            SQLCommand.Parameters.Clear()
            Dim oSQLParameter As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter
            If Not Equal(Parameters, Nothing) Then
                For iRowIndex As Integer = 0 To Parameters.GetLength(0) - 1
                    oSQLParameter = New System.Data.SqlClient.SqlParameter
                    oSQLParameter.ParameterName = "@" & Parameters(iRowIndex)(0)
                    Select Case Parameters(iRowIndex)(1)
                        Case DataTypes.Char
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Char
                            oSQLParameter.Size = Parameters(iRowIndex)(2)
                        Case DataTypes.DateTime
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.DateTime
                        Case DataTypes.Double
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Decimal
                        Case DataTypes.Image
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Image
                            oSQLParameter.Size = AdoConstants.MaxLargeDataLength
                        Case DataTypes.Integer
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Int
                        Case DataTypes.LargeText
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Text
                            oSQLParameter.Size = AdoConstants.MaxLargeDataLength
                        Case DataTypes.VarChar
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.VarChar
                            oSQLParameter.Size = Parameters(iRowIndex)(2)
                    End Select
                    oSQLParameter.Direction = ParameterDirection.Input
                    If Equal(Parameters(iRowIndex)(3), Nothing) Then oSQLParameter.Value = System.DBNull.Value Else oSQLParameter.Value = Parameters(iRowIndex)(3)
                    SQLCommand.Parameters.Add(oSQLParameter)
                Next
            Else
                ''Remove the openclose brackets in the procedure name
                If sProcedureName.EndsWith("()") Then SQLCommand.CommandText = sProcedureName.Substring(0, sProcedureName.Trim.Length - 2)
            End If
            SQLDataAdapter.Fill(dtDataTable)
        Catch ex As Exception
            If ex.Message.StartsWith("Validate : ") Then
                Throw ex
            Else
                Throw New Exception("Error occurred while executing report procedure - " & sProcedureName & vbCrLf & vbCrLf & ex.Message)
            End If
        Finally
            SQLCommand.Parameters.Clear()
        End Try
        Return dtDataTable
    End Function
    Public Function ExecuteProcedure(ByVal sProcedureName As String, ByVal Parameters()() As Object, ByRef dtDataTable As DataTable) As DataTable Implements IAdoDAL.ExecuteProcedure
        'Columns of array are: sParameterName, sDataType, iSize, oValue
        If dtDataTable Is Nothing Then dtDataTable = New DataTable
        Try
            SQLCommand.CommandText = sProcedureName
            SQLCommand.CommandType = CommandType.StoredProcedure
            SQLCommand.Transaction = SQLTransaction
            SQLCommand.Connection = SQLConnection
            SQLCommand.Parameters.Clear()
            Dim oSQLParameter As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter
            If Not Equal(Parameters, Nothing) Then
                For iRowIndex As Integer = 0 To Parameters.GetLength(0) - 1
                    oSQLParameter = New System.Data.SqlClient.SqlParameter
                    oSQLParameter.ParameterName = "@" & Parameters(iRowIndex)(0)
                    Select Case Parameters(iRowIndex)(1)
                        Case DataTypes.Char
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Char
                            oSQLParameter.Size = Parameters(iRowIndex)(2)
                        Case DataTypes.DateTime
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.DateTime
                        Case DataTypes.Double
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Decimal
                        Case DataTypes.Image
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Image
                            oSQLParameter.Size = AdoConstants.MaxLargeDataLength
                        Case DataTypes.Integer
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Int
                        Case DataTypes.LargeText
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Text
                            oSQLParameter.Size = AdoConstants.MaxLargeDataLength
                        Case DataTypes.VarChar
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.VarChar
                            oSQLParameter.Size = Parameters(iRowIndex)(2)
                    End Select
                    oSQLParameter.Direction = ParameterDirection.Input
                    If Equal(Parameters(iRowIndex)(3), Nothing) Then oSQLParameter.Value = System.DBNull.Value Else oSQLParameter.Value = Parameters(iRowIndex)(3)
                    SQLCommand.Parameters.Add(oSQLParameter)
                Next
            Else
                ''Remove the openclose brackets in the procedure name
                If sProcedureName.EndsWith("()") Then SQLCommand.CommandText = sProcedureName.Substring(0, sProcedureName.Trim.Length - 2)
            End If
            SQLDataAdapter.Fill(dtDataTable)
        Catch ex As Exception
            If ex.Message.StartsWith("Validate : ") Then
                Throw ex
            Else
                Throw New Exception("Error occurred while executing report procedure - " & sProcedureName & vbCrLf & vbCrLf & ex.Message)
            End If
        Finally
            SQLCommand.Parameters.Clear()
        End Try
        Return dtDataTable
    End Function
    ''Added overloaded feature for returning dataset 
    Public Function ExecuteProcedureDataSet(ByVal sProcedureName As String, ByVal Parameters()() As Object) As DataSet Implements IAdoDAL.ExecuteProcedureDataSet
        'Columns of array are: sParameterName, sDataType, iSize, oValue
        Dim dsDataSet As DataSet = New DataSet
        Try
            SQLCommand.CommandText = sProcedureName
            SQLCommand.CommandType = CommandType.StoredProcedure
            SQLCommand.Transaction = SQLTransaction
            SQLCommand.Connection = SQLConnection
            SQLCommand.Parameters.Clear()
            Dim oSQLParameter As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter
            If Not Equal(Parameters, Nothing) Then
                For iRowIndex As Integer = 0 To Parameters.GetLength(0) - 1
                    oSQLParameter = New System.Data.SqlClient.SqlParameter
                    oSQLParameter.ParameterName = "@" & Parameters(iRowIndex)(0)
                    Select Case Parameters(iRowIndex)(1)
                        Case DataTypes.Char
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Char
                            oSQLParameter.Size = Parameters(iRowIndex)(2)
                        Case DataTypes.DateTime
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.DateTime
                        Case DataTypes.Double
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Decimal
                        Case DataTypes.Image
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Image
                            oSQLParameter.Size = AdoConstants.MaxLargeDataLength
                        Case DataTypes.Integer
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Int
                        Case DataTypes.LargeText
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.Text
                            oSQLParameter.Size = AdoConstants.MaxLargeDataLength
                        Case DataTypes.VarChar
                            oSQLParameter.SqlDbType = System.Data.SqlDbType.VarChar
                            oSQLParameter.Size = Parameters(iRowIndex)(2)
                    End Select
                    oSQLParameter.Direction = ParameterDirection.Input
                    If Equal(Parameters(iRowIndex)(3), Nothing) Then oSQLParameter.Value = System.DBNull.Value Else oSQLParameter.Value = Parameters(iRowIndex)(3)
                    SQLCommand.Parameters.Add(oSQLParameter)
                Next
            End If
            SQLDataAdapter.Fill(dsDataSet)
        Catch ex As Exception
            'Dim sCustomError As String = ""
            'For iRowIndex As Integer = 0 To Parameters.GetLength(0) - 1
            '    If sCustomError <> "" Then sCustomError += ", "
            '    If Equal(Parameters(iRowIndex)(3), Nothing) Then
            '        sCustomError += Parameters(iRowIndex)(0) + ":Nothing"
            '    Else
            '        sCustomError += Parameters(iRowIndex)(0) + ":" + Parameters(iRowIndex)(3).ToString
            '    End If
            'Next
            'If sCustomError <> "" Then sCustomError = "Parameter:- " + sCustomError
            LogSystemError(ex, "ProcedureName : " & sProcedureName)
            Throw New Exception("Error occurred while executing report procedure - " & sProcedureName & vbCrLf & vbCrLf & ex.Message)
        Finally
            SQLCommand.Parameters.Clear()
        End Try
        Return dsDataSet
    End Function

    Public Sub BeginTransaction() Implements IAdoDAL.BeginTransaction
        Try
            oSQLTransaction = oSQLConnection.BeginTransaction(IsolationLevel.ReadCommitted)
            bInTransaction = True
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
    End Sub
    Public Sub Commit() Implements IAdoDAL.Commit
        Try
            SQLTransaction.Commit()
            oSQLTransaction.Dispose()
            oSQLTransaction = Nothing
            bInTransaction = False
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
    End Sub
    Public Sub RollBack() Implements IAdoDAL.RollBack
        Try
            SQLTransaction.Rollback()
            oSQLTransaction.Dispose()
            oSQLTransaction = Nothing
            bInTransaction = False
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
    End Sub
    Public Function UpdateRows(ByVal oUpdateCommand As Data.IDbCommand, ByVal oDataRows As DataRow()) As Integer Implements IAdoDAL.UpdateRows
        Dim iNumberOfRowsUpdated As Integer = 0
        Try
            SQLDataAdapter.InsertCommand = oUpdateCommand
            SQLDataAdapter.UpdateCommand = oUpdateCommand
            SQLDataAdapter.UpdateCommand.Connection = SQLConnection
            SQLDataAdapter.UpdateCommand.Transaction = SQLTransaction
            iNumberOfRowsUpdated = SQLDataAdapter.Update(oDataRows)
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        Return iNumberOfRowsUpdated
    End Function
    Public ReadOnly Property ConnectionState() As System.Data.ConnectionState Implements IAdoDAL.ConnectionState
        Get
            Return oSQLConnection.State
        End Get
    End Property
    Public ReadOnly Property InTransaction() As Boolean Implements IAdoDAL.InTransaction
        Get
            Return bInTransaction
        End Get
    End Property

    Private Function BuildDataTable(ByVal oSQLReader As IDataReader, ByRef ResultsData As DataTable) As List(Of DataColumn)
        Dim dtSchema As DataTable = oSQLReader.GetSchemaTable()
        Dim listCols = New List(Of DataColumn)
        If dtSchema IsNot Nothing Then
            For Each drow As DataRow In dtSchema.Rows
                Dim columnName As String = Convert.ToString(drow("ColumnName"))
                Dim column = New DataColumn(columnName, DirectCast(drow("DataType"), Type))
                column.Unique = CBool(drow("IsUnique"))
                column.AllowDBNull = CBool(drow("AllowDBNull"))
                column.AutoIncrement = CBool(drow("IsAutoIncrement"))
                listCols.Add(column)
                ResultsData.Columns.Add(column)
            Next
        End If
        Return listCols
    End Function

    Private Sub ReadSQLDataReader(oSQLReader As SqlDataReader, ResultsData As DataTable, iRecordLimit As Integer, listCols As List(Of DataColumn))
        Dim iIndex As Integer = 0
        Dim firstTime As Boolean = True
        Try
            While oSQLReader.Read()
                Dim dataRow As DataRow = ResultsData.NewRow()
                For i As Integer = 0 To listCols.Count - 1
                    dataRow((listCols(i))) = oSQLReader(i)
                Next
                ResultsData.Rows.Add(dataRow)
                iIndex += 1
                If iIndex = iRecordLimit Then
                    iIndex = 0
                    ''ExportToOxml(firstTime)
                    'RaiseEvent ExportToXML(firstTime, ResultsData)
                    PartialData(firstTime, ResultsData)
                    ResultsData.Clear()
                    firstTime = False
                End If
            End While
            If ResultsData.Rows.Count > 0 Then
                ''ExportToOxml(firstTime)
                ''RaiseEvent ExportToXML(firstTime, ResultsData)
                PartialData(firstTime, ResultsData)
                ResultsData.Clear()
            End If
            ' Call Close when done reading.
            oSQLReader.Close()
        Catch ex As Exception
            oSQLReader.Close()
            LogSystemError(ex)
            Throw ex
        End Try
    End Sub
   

    Public Function BulkCopy(dtSource As System.Data.DataTable, sDestination As String) As Boolean Implements AdoDALInterface.IAdoDAL.BulkCopy
        Dim bSuccess As Boolean
        If dtSource.Rows.Count > 0 Then
            ' Code for bulk data writing 
            Dim sbc As SqlBulkCopy = New SqlBulkCopy(SQLConnection)
            Try
                'Initializing an SqlBulkCopy object
                sbc.DestinationTableName = sDestination
                sbc.WriteToServer(dtSource)
                sbc.Close()
                bSuccess = True
            Catch ex As Exception
                bSuccess = False
                sbc.Close()
                LogSystemError(ex)
                Throw ex
            End Try
        End If
        Return bSuccess
    End Function
End Class
