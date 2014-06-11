Imports AdoDALInterface

Public Class DataBaseInformation
    Implements IDataBaseInfo

    Public Property ConnectionID As String Implements AdoDALInterface.IDataBaseInfo.ConnectionID
    Public Property DatabaseName As String Implements AdoDALInterface.IDataBaseInfo.DatabaseName
    Public Property DatabasePassword As String Implements AdoDALInterface.IDataBaseInfo.DatabasePassword
    Public Property DatabasePort As String Implements AdoDALInterface.IDataBaseInfo.DatabasePort
    Public Property DatabaseServerName As String Implements AdoDALInterface.IDataBaseInfo.DatabaseServerName
    Public Property DatabaseType As String Implements AdoDALInterface.IDataBaseInfo.DatabaseType
    Public Property DatabaseTypeName As String Implements AdoDALInterface.IDataBaseInfo.DatabaseTypeName
    Public Property DatabaseUserName As String Implements AdoDALInterface.IDataBaseInfo.DatabaseUserName
    Public Property Connection As System.Data.Common.DbConnection Implements AdoDALInterface.IDataBaseInfo.Connection
End Class
