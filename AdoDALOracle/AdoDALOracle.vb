Imports Utilities.Utilities
Imports AdoDALInterface
Imports AdoDALInterface.AdoConstants

Public Class AdoDALOracle
    Implements IAdoDAL



    Private oOracleConnection As System.Data.Common.DbConnection = Nothing '' Oracle.DataAccess.Client.OracleConnection = Nothing
    Private oOracleCommand As System.Data.Common.DbCommand = Nothing ''Oracle.DataAccess.Client.OracleCommand = Nothing
    Private oOracleDataAdapter As IDbDataAdapter = Nothing '' Oracle.DataAccess.Client.OracleDataAdapter = Nothing
    Private oOracleTransaction As IDbTransaction = Nothing '' Oracle.DataAccess.Client.OracleTransaction = Nothing
    Private bInTransaction As Boolean = False
    'Private CurrentDAL As AdoDALOracle = Me

    Public Property DatabaseInfo As AdoDALInterface.IDataBaseInfo Implements AdoDALInterface.IAdoDAL.DatabaseInfo
    Public Property DataBaseErrorMessages As List(Of String) = New List(Of String) Implements IAdoDAL.DataBaseErrorMessages
    Public Property DataBaseMessages As List(Of String) Implements IAdoDAL.DataBaseMessages
    Public Overridable Sub PartialData(ByVal firstTime As Boolean, ResultsData As DataTable)

    End Sub

    Private ReadOnly Property OracleConnection() As Oracle.DataAccess.Client.OracleConnection
        Get
            If Equal(oOracleCommand, Nothing) OrElse Equal(oOracleConnection, Nothing) OrElse oOracleConnection.State = ConnectionState.Broken OrElse oOracleConnection.State = ConnectionState.Closed Then
                If Equal(oOracleTransaction, Nothing) Then Connect() Else DataBaseErrorMessages.Add("OracleConnection is lost while saving. Please try saving again")
            End If
            Return oOracleConnection
        End Get
    End Property
    Private ReadOnly Property OracleCommand() As Oracle.DataAccess.Client.OracleCommand
        Get
            If Equal(oOracleCommand, Nothing) OrElse Equal(oOracleConnection, Nothing) OrElse oOracleConnection.State = ConnectionState.Broken OrElse oOracleConnection.State = ConnectionState.Closed Then
                If Equal(oOracleTransaction, Nothing) Then Connect() Else DataBaseErrorMessages.Add("OracleConnection is lost while saving. Please try saving again")
            End If
            Return oOracleCommand
        End Get
    End Property
    Private ReadOnly Property OracleDataAdapter() As Oracle.DataAccess.Client.OracleDataAdapter
        Get
            If Equal(oOracleCommand, Nothing) OrElse Equal(oOracleConnection, Nothing) OrElse oOracleConnection.State = ConnectionState.Broken OrElse oOracleConnection.State = ConnectionState.Closed Then
                If Equal(oOracleTransaction, Nothing) Then Connect() Else DataBaseErrorMessages.Add("OracleConnection is lost while saving. Please try saving again")
            End If
            Return oOracleDataAdapter
        End Get
    End Property
    Private ReadOnly Property OracleTransaction() As Oracle.DataAccess.Client.OracleTransaction
        Get
            If Equal(oOracleCommand, Nothing) OrElse Equal(oOracleConnection, Nothing) OrElse oOracleConnection.State = ConnectionState.Broken OrElse oOracleConnection.State = ConnectionState.Closed Then
                If Equal(oOracleTransaction, Nothing) Then Connect() Else DataBaseErrorMessages.Add("OracleConnection is lost while saving. Please try saving again")
            End If
            Return oOracleTransaction
        End Get
    End Property
    Public ReadOnly Property ConnectionState() As System.Data.ConnectionState Implements IAdoDAL.ConnectionState
        Get
            Return oOracleConnection.State
        End Get
    End Property
    Public ReadOnly Property InTransaction() As Boolean Implements IAdoDAL.InTransaction
        Get
            Return bInTransaction
        End Get
    End Property
    Public Overridable Function ExecuteQuery(sQuery As String) As System.Data.DataTable Implements IAdoDAL.ExecuteQuery
        Dim dtDataTable As DataTable = New DataTable
        Dim ds As DataSet = New DataSet
        Try
            sQuery = sQuery.Replace("dbo.", "")
            OracleCommand.CommandText = sQuery
            OracleCommand.CommandType = CommandType.Text
            OracleCommand.Connection = OracleConnection
            ds.Tables.Add(dtDataTable)
            OracleDataAdapter.Fill(dtDataTable)
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        Return dtDataTable ''ds.Tables(0) ''
    End Function
    Public Overridable Sub ExecuteQuery(ByVal sQuery As String, ByVal iRecordLimit As Integer) Implements IAdoDAL.ExecuteQuery 'As DataTable 
        Dim dtDataTable As DataTable = New DataTable
        Try
            OracleCommand.CommandText = sQuery
            OracleCommand.CommandType = CommandType.Text
            OracleCommand.Connection = OracleConnection

            Dim oSQLReader As IDataReader = OracleCommand.ExecuteReader()
            Dim ListOfColumn As List(Of DataColumn) = BuildDataTable(oSQLReader, dtDataTable)
            ''CallReader 
            ReadSQLDataReader(oSQLReader, dtDataTable, iRecordLimit, ListOfColumn)
            ''End If
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        ''Return dtDataTable
    End Sub
    Public Function ExecuteScalar(ByVal sQuery As String) As Object Implements IAdoDAL.ExecuteScalar
        Dim oObject As Object = Nothing
        Try
            sQuery = sQuery.Replace("dbo.", "")
            OracleCommand.CommandText = sQuery
            OracleCommand.CommandType = CommandType.Text
            OracleCommand.Connection = OracleConnection
            oObject = OracleCommand.ExecuteScalar
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
    Public Function ExecuteNonQuery(ByVal sQuery As String) As Integer Implements IAdoDAL.ExecuteNonQuery
        Dim iRowsAffected As Integer = 0
        Try
            sQuery = sQuery.Replace("dbo.", "")
            OracleCommand.CommandText = sQuery
            OracleCommand.CommandType = CommandType.Text
            OracleCommand.Connection = OracleConnection
            iRowsAffected = OracleCommand.ExecuteNonQuery
        Catch ex As Exception
            'LogSystemError(ex)
            Throw ex
        End Try
        Return iRowsAffected
    End Function
    Public Function ExecuteProcedure(ByVal sProcedureName As String, ByVal Parameters()() As Object) As DataTable Implements IAdoDAL.ExecuteProcedure
        'Columns of array are: sParameterName, sDataType, iSize, oValue
        Dim dtDataTable As DataTable = New DataTable
        Try
            OracleCommand.CommandText = sProcedureName
            OracleCommand.CommandType = CommandType.StoredProcedure
            OracleCommand.Connection = OracleConnection
            OracleCommand.Parameters.Clear()
            Dim oOracleParameter As Oracle.DataAccess.Client.OracleParameter = New Oracle.DataAccess.Client.OracleParameter
            If Not Equal(Parameters, Nothing) Then
                For iRowIndex As Integer = 0 To Parameters.GetLength(0) - 1
                    oOracleParameter = New Oracle.DataAccess.Client.OracleParameter
                    oOracleParameter.ParameterName = Parameters(iRowIndex)(0)
                    Select Case Parameters(iRowIndex)(1)
                        Case DataTypes.Char
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Char
                        Case DataTypes.DateTime
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Date
                        Case DataTypes.Double
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Decimal
                        Case DataTypes.Image
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Blob
                        Case DataTypes.Integer
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Int32
                        Case DataTypes.LargeText
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Clob
                        Case DataTypes.VarChar
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Varchar2
                    End Select
                    oOracleParameter.Direction = ParameterDirection.Input
                    If Equal(Parameters(iRowIndex)(3), Nothing) Then oOracleParameter.Value = System.DBNull.Value Else oOracleParameter.Value = Parameters(iRowIndex)(3)
                    OracleCommand.Parameters.Add(oOracleParameter)
                Next
            End If
            oOracleParameter = New Oracle.DataAccess.Client.OracleParameter
            oOracleParameter.ParameterName = "ResultCursor"
            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.RefCursor
            oOracleParameter.Direction = ParameterDirection.Output
            OracleCommand.Parameters.Add(oOracleParameter)

            OracleDataAdapter.Fill(dtDataTable)
        Catch ex As Exception
            Throw New Exception("Error occurred while executing report procedure - " & sProcedureName & vbCrLf & vbCrLf & ex.Message)
        Finally
            OracleCommand.Parameters.Clear()
        End Try
        Return dtDataTable
    End Function
    Public Function ExecuteProcedure(ByVal sProcedureName As String, ByVal Parameters()() As Object, ByRef dtDataTable As DataTable) As DataTable Implements IAdoDAL.ExecuteProcedure
        'Columns of array are: sParameterName, sDataType, iSize, oValue
        If dtDataTable Is Nothing Then dtDataTable = New DataTable
        Try
            OracleCommand.CommandText = sProcedureName
            OracleCommand.CommandType = CommandType.StoredProcedure
            OracleCommand.Connection = OracleConnection
            OracleCommand.Parameters.Clear()
            Dim oOracleParameter As Oracle.DataAccess.Client.OracleParameter = New Oracle.DataAccess.Client.OracleParameter
            If Not Equal(Parameters, Nothing) Then
                For iRowIndex As Integer = 0 To Parameters.GetLength(0) - 1
                    oOracleParameter = New Oracle.DataAccess.Client.OracleParameter
                    oOracleParameter.ParameterName = Parameters(iRowIndex)(0)
                    Select Case Parameters(iRowIndex)(1)
                        Case DataTypes.Char
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Char
                        Case DataTypes.DateTime
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Date
                        Case DataTypes.Double
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Decimal
                        Case DataTypes.Image
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Blob
                        Case DataTypes.Integer
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Int32
                        Case DataTypes.LargeText
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Clob
                        Case DataTypes.VarChar
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Varchar2
                    End Select
                    oOracleParameter.Direction = ParameterDirection.Input
                    If Equal(Parameters(iRowIndex)(3), Nothing) Then oOracleParameter.Value = System.DBNull.Value Else oOracleParameter.Value = Parameters(iRowIndex)(3)
                    OracleCommand.Parameters.Add(oOracleParameter)
                Next
            End If
            oOracleParameter = New Oracle.DataAccess.Client.OracleParameter
            oOracleParameter.ParameterName = "ResultCursor"
            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.RefCursor
            oOracleParameter.Direction = ParameterDirection.Output
            OracleCommand.Parameters.Add(oOracleParameter)

            OracleDataAdapter.Fill(dtDataTable)
        Catch ex As Exception
            Throw New Exception("Error occurred while executing report procedure - " & sProcedureName & vbCrLf & vbCrLf & ex.Message)
        Finally
            OracleCommand.Parameters.Clear()
        End Try
        Return dtDataTable
    End Function
    Public Function ExecuteProcedureDataSet(ByVal sProcedureName As String, ByVal Parameters()() As Object) As DataSet Implements IAdoDAL.ExecuteProcedureDataSet
        'Columns of array are: sParameterName, sDataType, iSize, oValue
        Dim dsDataSet As DataSet = New DataSet
        Try
            'Retrieve all the parameter of the procedure which is of type RefCursor and store it in dtProcedureArguments
            Dim dtProcedureArguments As DataTable = Nothing
            dtProcedureArguments = Me.ExecuteQuery("SELECT Argument_Name, Position, Data_Type FROM ALL_ARGUMENTS WHERE OBJECT_NAME = '" & sProcedureName.Trim.ToUpper & "' AND Data_Type = 'REF CURSOR' AND OWNER = USER ORDER BY Position")

            OracleCommand.CommandText = sProcedureName
            OracleCommand.CommandType = CommandType.StoredProcedure
            OracleCommand.Connection = OracleConnection
            OracleCommand.Parameters.Clear()
            Dim oOracleParameter As Oracle.DataAccess.Client.OracleParameter = New Oracle.DataAccess.Client.OracleParameter
            If Not Equal(Parameters, Nothing) Then
                For iRowIndex As Integer = 0 To Parameters.GetLength(0) - 1
                    oOracleParameter = New Oracle.DataAccess.Client.OracleParameter
                    oOracleParameter.ParameterName = Parameters(iRowIndex)(0)
                    Select Case Parameters(iRowIndex)(1)
                        Case DataTypes.Char
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Char
                            oOracleParameter.Size = Parameters(iRowIndex)(2)
                        Case DataTypes.DateTime
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Date
                        Case DataTypes.Double
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Decimal
                        Case DataTypes.Image
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Blob
                            oOracleParameter.Size = AdoConstants.MaxLargeDataLength
                        Case DataTypes.Integer
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Int32
                        Case DataTypes.LargeText
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Clob
                            oOracleParameter.Size = AdoConstants.MaxLargeDataLength
                        Case DataTypes.VarChar
                            oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.Varchar2
                            oOracleParameter.Size = Parameters(iRowIndex)(2)
                    End Select
                    oOracleParameter.Direction = ParameterDirection.Input
                    If Equal(Parameters(iRowIndex)(3), Nothing) Then oOracleParameter.Value = System.DBNull.Value Else oOracleParameter.Value = Parameters(iRowIndex)(3)
                    OracleCommand.Parameters.Add(oOracleParameter)
                Next
            End If

            For Each drRow As DataRow In dtProcedureArguments.Rows
                oOracleParameter = New Oracle.DataAccess.Client.OracleParameter
                oOracleParameter.ParameterName = drRow("Argument_Name")
                oOracleParameter.OracleDbType = Oracle.DataAccess.Client.OracleDbType.RefCursor
                oOracleParameter.Direction = ParameterDirection.Output
                OracleCommand.Parameters.Add(oOracleParameter)
            Next

            OracleDataAdapter.Fill(dsDataSet)
        Catch ex As Exception
            Throw ex
        Finally
            OracleCommand.Parameters.Clear()
        End Try
        Return dsDataSet
    End Function
    Public Function Connect() As System.Data.Common.DbConnection Implements IAdoDAL.Connect
        Try

            Dim sPassword As String = String.Empty
            If Not Equal(DatabaseInfo.DatabasePassword, Nothing) Then
                If DatabaseInfo.DatabasePassword.IndexOf("~") > 0 Then ''Changed for Web
                    sPassword = Decrypt(DatabaseInfo.DatabasePassword)
                Else
                    sPassword = DatabaseInfo.DatabasePassword
                End If
            End If

            Dim sConnectionString As String = "Data Source=(DESCRIPTION=" _
                       + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & DatabaseInfo.DatabaseServerName & ")(PORT=" & DatabaseInfo.DatabasePort & ")))" _
                       + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=" & DatabaseInfo.DatabaseName & ")));" _
                       + "User Id=" & DatabaseInfo.DatabaseUserName & ";Password=" & sPassword & ";"

            oOracleConnection = New Oracle.DataAccess.Client.OracleConnection(sConnectionString)
            Dim sCommandText As String = Nothing
            sCommandText = "ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY HH24:MI:SS'"
            Try
                oOracleConnection.Open()
            Catch ex As Exception
                DataBaseErrorMessages.Add("Unable to establish database connection." & vbCrLf & "Please check database configuration.")
                ''Base.BaseApplication.oApplication.ExitApplication()
                'MsgBox(ex.Message, MsgBoxButtons.OK)
            End Try
            If oOracleConnection.State = ConnectionState.Open Then
                oOracleCommand = New Oracle.DataAccess.Client.OracleCommand(sCommandText, oOracleConnection)
                oOracleCommand.CommandTimeout = 0  'Changes done by Roshan.
                oOracleDataAdapter = New Oracle.DataAccess.Client.OracleDataAdapter(oOracleCommand)
                oOracleCommand.ExecuteNonQuery()
                DatabaseInfo.ConnectionID = Me.ExecuteScalar("SELECT 'pid ' || SUBSTR(a.spid,1,9) || ' - ser# ' || SUBSTR(b.serial#,1,5) PID FROM v$session b, v$process a WHERE b.paddr = a.addr AND type='USER' ORDER BY spid")
            End If
            DatabaseInfo.Connection = OracleConnection
        Catch ex As Exception

        End Try
        Return oOracleConnection
    End Function
    Public Function Connect(sApplicationName As String) As IDbConnection Implements IAdoDAL.Connect
        Return Connect()
    End Function
    Public Function Close() As Boolean Implements IAdoDAL.Close
        Dim bSuccess As Boolean = False
        Try
            If oOracleConnection.State = ConnectionState.Open Then oOracleConnection.Close()
            bSuccess = True
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        Return bSuccess
    End Function
    Public Sub BeginTransaction() Implements IAdoDAL.BeginTransaction
        Try
            oOracleTransaction = oOracleConnection.BeginTransaction(IsolationLevel.ReadCommitted)
            bInTransaction = True
        Catch ex As Exception
            LogSystemError(ex)
        End Try
    End Sub
    Public Sub Commit() Implements IAdoDAL.Commit
        Try
            OracleTransaction.Commit()
            oOracleTransaction.Dispose()
            oOracleTransaction = Nothing
            bInTransaction = False
        Catch ex As Exception
            LogSystemError(ex)
        End Try
    End Sub
    Public Sub RollBack() Implements IAdoDAL.RollBack
        Try
            OracleTransaction.Rollback()
            oOracleTransaction.Dispose()
            oOracleTransaction = Nothing
            bInTransaction = False
        Catch ex As Exception
            LogSystemError(ex)
        End Try
    End Sub
    Public Function UpdateRows(ByVal oUpdateCommand As Data.IDbCommand, ByVal oDataRows As DataRow()) As Integer Implements IAdoDAL.UpdateRows
        Dim iNumberOfRowsUpdated As Integer = 0
        Try
            OracleDataAdapter.InsertCommand = oUpdateCommand
            OracleDataAdapter.UpdateCommand = oUpdateCommand
            OracleDataAdapter.UpdateCommand.Connection = OracleConnection
            'OracleDataAdapter.UpdateCommand.Transaction = OracleTransaction
            iNumberOfRowsUpdated = OracleDataAdapter.Update(oDataRows)
        Catch ex As Exception
            LogSystemError(ex)
            Throw ex
        End Try
        Return iNumberOfRowsUpdated
    End Function
    Public Overridable Function ExecuteNonQueryAsync(sQuery As String) As Integer Implements IAdoDAL.ExecuteNonQueryAsync
        ''TODO: implementation
        Return Nothing
    End Function
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

    Private Sub ReadSQLDataReader(oSQLReader As IDataReader, ResultsData As DataTable, iRecordLimit As Integer, listCols As List(Of DataColumn))
        Dim iIndex As Integer = 0
        Dim firstTime As Boolean = True
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
    End Sub

    Public Function BulkCopy(dtSource As System.Data.DataTable, sDestination As String) As Boolean Implements AdoDALInterface.IAdoDAL.BulkCopy
        Return False
    End Function
End Class
