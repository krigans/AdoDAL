Imports AdoDALInterface

Public Class DataBaseAdapter
    Implements IDataBaseAdapter

    Public Property Connection As IDbConnection Implements IDataBaseAdapter.Connection
    Public Property Command As IDbCommand Implements IDataBaseAdapter.Command
    Public Property DataAdapter As IDataAdapter Implements IDataBaseAdapter.DataAdapter
    Public Property Transaction As IDbTransaction Implements IDataBaseAdapter.Transaction
    Public Property DataBaseInfo As AdoDALInterface.IDataBaseInfo Implements AdoDALInterface.IDataBaseAdapter.DataBaseInfo
End Class
