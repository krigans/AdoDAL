Public Interface IDataBaseAdapter
    Property Connection As IDbConnection
    Property Command As IDbCommand
    Property DataAdapter As IDataAdapter
    Property Transaction As IDbTransaction
    Property DataBaseInfo As IDataBaseInfo
End Interface
