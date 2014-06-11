Imports StructureMap
Imports AdoDALInterface

Public Class StructureMapBootStrapper
    Private Sub New()
    End Sub

    'For debugging uncomment the below and use it
    'Public Shared Sub BootStrap()
    '    StructureMap.ObjectFactory.Initialize(Sub(Instance)
    '                                              Instance.[For](Of IAdoDAL)().Add(Of AdoDALSQL.AdoDALSQL)().Named("MSSQL")
    '                                              Instance.[For](Of IAdoDAL)().Add(Of AdoDALOracle.AdoDALOracle)().Named("ORACLE")
    '                                          End Sub)
    'End Sub

    Public Shared Sub BootStrap()
        Try
            StructureMap.ObjectFactory.Initialize(Sub(Instance)
                                                      Try
                                                          Instance.PullConfigurationFromAppConfig = True
                                                      Catch ex As Exception

                                                      End Try
                                                  End Sub)
        Catch ex As Exception
            Console.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    ''For Automatic Loading
    ''Public Shared Sub BootStrap()
    ''    ObjectFactory.Initialize(Function(Registry)
    ''                                 Registry.Scan(Function(assembly)
    ''                                                   assembly.TheCallingAssembly()

    ''                                                   'Telling StructureMap to sweep a folder called "extensions" directly
    ''                                                   'underneath the application root folder for any assemblies found in that folder
    ''                                                   assembly.AssembliesFromPath("DBExtensions", Function(addedAssembly) addedAssembly.GetName().Name.ToLower().Contains("AdoDAL"))

    ''                                                   'Direct StructureMap to add any Registries that it finds in these assemblies, assuming that all the StructureMap directives are
    ''                                                   'contained in registry classes
    ''                                                   ''assembly.LookForRegistries()

    ''                                                   assembly.AddAllTypesOf(Of IAdoDAL)()
    ''                                               End Function)

    ''                             End Function)
    ''End Sub

End Class
