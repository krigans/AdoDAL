Imports AdoDALInterface
Imports AdoDALInterface.AdoConstants

Public Class AdoDAL
    Implements IAdoDAL
    Implements ICurrentDAL

    Private Property oCurrentDAL As IAdoDAL

    Public Property DatabaseInfo As AdoDALInterface.IDataBaseInfo Implements AdoDALInterface.IAdoDAL.DatabaseInfo
    Public Property DataBaseErrorMessages As List(Of String) Implements IAdoDAL.DataBaseErrorMessages
    Public Property DataBaseMessages As List(Of String) Implements IAdoDAL.DataBaseMessages

    Shared Sub New()
        StructureMapBootStrapper.BootStrap()
    End Sub
    Private Sub New()
        ReadConfig()
        SetDBConfigs()
    End Sub
    Public Shared Function GetInstance() As IAdoDAL
        Return (New AdoDAL).CurrentDAL
    End Function
    Public Sub ReadConfig()
        Try
            'Read the config file and store the config details in DAL
            DatabaseInfo = New DataBaseInformation
            DatabaseInfo.DatabaseTypeName = System.Configuration.ConfigurationManager.AppSettings("DatabaseType")
            Me.DatabaseInfo.DatabaseServerName = System.Configuration.ConfigurationManager.AppSettings("DatabaseServerName")
            Me.DatabaseInfo.DatabasePort = System.Configuration.ConfigurationManager.AppSettings("DatabasePort")
            Me.DatabaseInfo.DatabaseName = System.Configuration.ConfigurationManager.AppSettings("DatabaseName")
            Me.DatabaseInfo.DatabaseUserName = System.Configuration.ConfigurationManager.AppSettings("DatabaseUserName")
            Me.DatabaseInfo.DatabasePassword = System.Configuration.ConfigurationManager.AppSettings("DatabasePassword")
            oCurrentDAL = StructureMap.ObjectFactory.GetNamedInstance(Of IAdoDAL)(Me.DatabaseInfo.DatabaseTypeName)
            Select Case Me.DatabaseInfo.DatabaseTypeName
                Case "ORACLE"
                    Me.DatabaseInfo.DatabaseType = DatabaseTypes.Oracle
                Case "MSSQL"
                    Me.DatabaseInfo.DatabaseType = DatabaseTypes.MSSql
            End Select
            'Me.DatabaseInfo.Connection = oCurrentDAL.DatabaseInfo.Connection
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub
    Private Sub SetDBConfigs()
        Try
            If oCurrentDAL.DatabaseInfo Is Nothing Then oCurrentDAL.DatabaseInfo = New DataBaseInformation
            oCurrentDAL.DatabaseInfo.DatabaseServerName = Me.DatabaseInfo.DatabaseServerName
            oCurrentDAL.DatabaseInfo.DatabasePort = Me.DatabaseInfo.DatabasePort
            oCurrentDAL.DatabaseInfo.DatabaseName = Me.DatabaseInfo.DatabaseName
            oCurrentDAL.DatabaseInfo.DatabaseUserName = Me.DatabaseInfo.DatabaseUserName
            oCurrentDAL.DatabaseInfo.DatabasePassword = Me.DatabaseInfo.DatabasePassword
            oCurrentDAL.DatabaseInfo.Connection = Me.DatabaseInfo.Connection
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Public ReadOnly Property CurrentDAL As AdoDALInterface.IAdoDAL Implements AdoDALInterface.ICurrentDAL.CurrentDAL
        Get
            If oCurrentDAL Is Nothing Then
                ReadConfig()
                SetDBConfigs()
            End If

            Return oCurrentDAL
        End Get
    End Property
    Public Overridable Function ExecuteQuery(sQuery As String) As System.Data.DataTable Implements IAdoDAL.ExecuteQuery
        Return oCurrentDAL.ExecuteQuery(sQuery)
    End Function
    Public Function ExecuteScalar(sQuery As String) As Object Implements AdoDALInterface.IAdoDAL.ExecuteScalar
        Return oCurrentDAL.ExecuteQuery(sQuery)
    End Function
    Public Function Close() As Boolean Implements AdoDALInterface.IAdoDAL.Close
        Return oCurrentDAL.Close
    End Function
    Public Function Connect() As System.Data.Common.DbConnection Implements AdoDALInterface.IAdoDAL.Connect
        Return oCurrentDAL.Connect
    End Function
    Public Function Connect(sApplicationName As String) As System.Data.IDbConnection Implements AdoDALInterface.IAdoDAL.Connect
        Return oCurrentDAL.Connect(sApplicationName)
    End Function
    Public Sub BeginTransaction() Implements IAdoDAL.BeginTransaction
        oCurrentDAL.BeginTransaction()
    End Sub
    Public Sub Commit() Implements IAdoDAL.Commit
        oCurrentDAL.Commit()
    End Sub
    Public Sub RollBack() Implements IAdoDAL.RollBack
        oCurrentDAL.RollBack()
    End Sub
    Public Function UpdateRows(ByVal oUpdateCommand As Data.IDbCommand, ByVal oDataRows As DataRow()) As Integer Implements IAdoDAL.UpdateRows
        Return oCurrentDAL.UpdateRows(oUpdateCommand, oDataRows)
    End Function
    Public ReadOnly Property ConnectionState As System.Data.ConnectionState Implements AdoDALInterface.IAdoDAL.ConnectionState
        Get
            Return oCurrentDAL.ConnectionState
        End Get
    End Property
    Public Function ExecuteNonQuery(sQuery As String) As Integer Implements AdoDALInterface.IAdoDAL.ExecuteNonQuery
        Return oCurrentDAL.ExecuteNonQuery(sQuery)
    End Function
    Public Function ExecuteProcedure(sProcedureName As String, Parameters()() As Object) As System.Data.DataTable Implements AdoDALInterface.IAdoDAL.ExecuteProcedure
        Return oCurrentDAL.ExecuteProcedure(sProcedureName, Parameters)
    End Function
    Public Function ExecuteProcedure(sProcedureName As String, Parameters()() As Object, ByRef dtDataTable As System.Data.DataTable) As System.Data.DataTable Implements AdoDALInterface.IAdoDAL.ExecuteProcedure
        Return oCurrentDAL.ExecuteProcedure(sProcedureName, Parameters)
    End Function
    Public Function ExecuteProcedureDataSet(sProcedureName As String, Parameters()() As Object) As System.Data.DataSet Implements AdoDALInterface.IAdoDAL.ExecuteProcedureDataSet
        Return oCurrentDAL.ExecuteProcedureDataSet(sProcedureName, Parameters)
    End Function
    Public ReadOnly Property InTransaction As Boolean Implements AdoDALInterface.IAdoDAL.InTransaction
        Get
            Return oCurrentDAL.InTransaction
        End Get
    End Property
    Public Overridable Function ExecuteNonQueryAsync(sQuery As String) As Integer Implements IAdoDAL.ExecuteNonQueryAsync
        Return oCurrentDAL.ExecuteNonQueryAsync(sQuery)
    End Function

    Public Overridable Sub ExecuteQuery(ByVal sQuery As String, ByVal iRecordLimit As Integer) Implements IAdoDAL.ExecuteQuery 'As DataTable 
        oCurrentDAL.ExecuteQuery(sQuery, iRecordLimit)
    End Sub

    Public Overridable Function BulkCopy(ByVal dtSource As DataTable, ByVal sDestination As String) As Boolean Implements IAdoDAL.BulkCopy
        Return oCurrentDAL.BulkCopy(dtSource, sDestination)
    End Function
End Class