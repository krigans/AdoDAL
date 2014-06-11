Public Interface IAdoDAL
    Property DatabaseInfo As IDataBaseInfo
    Property DataBaseErrorMessages As List(Of String)
    Property DataBaseMessages As List(Of String)
    ReadOnly Property ConnectionState() As System.Data.ConnectionState
    ReadOnly Property InTransaction() As Boolean

    Function Close() As Boolean
    Function Connect() As System.Data.Common.DbConnection ''IDbConnection
    Function Connect(sApplicationName As String) As IDbConnection
    Function ExecuteQuery(ByVal sQuery As String) As DataTable
    Sub ExecuteQuery(ByVal sQuery As String, ByVal iRecordLimit As Integer)
    Function ExecuteScalar(ByVal sQuery As String) As Object
    Function UpdateRows(ByVal oUpdateCommand As Data.IDbCommand, ByVal oDataRows As DataRow()) As Integer
    Function ExecuteNonQuery(ByVal sQuery As String) As Integer
    Function ExecuteNonQueryAsync(ByVal sQuery As String) As Integer
    Function ExecuteProcedure(ByVal sProcedureName As String, ByVal Parameters()() As Object) As DataTable
    Function ExecuteProcedure(ByVal sProcedureName As String, ByVal Parameters()() As Object, ByRef dtDataTable As DataTable) As DataTable
    Function ExecuteProcedureDataSet(ByVal sProcedureName As String, ByVal Parameters()() As Object) As DataSet
    Function BulkCopy(ByVal dtSource As DataTable, ByVal sDestination As String) As Boolean

    Sub BeginTransaction()
    Sub Commit()
    Sub RollBack()
End Interface
