Imports StructureMap.Configuration.DSL
Imports AdoDALInterface

Public Class DalRegistry
    Inherits Registry
    Public Sub New()
        Scan(Sub(x)
                 x.TheCallingAssembly()
                 x.AddAllTypesOf(Of IAdoDAL)().NameBy(Function(t) t.Name.ToUpper())
             End Sub)
    End Sub
End Class
