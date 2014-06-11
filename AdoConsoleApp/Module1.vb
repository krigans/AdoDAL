Imports AdoDALInterface

Module Module1

    Sub Main()
        Try
            Dim oAdo As IAdoDAL = AdoDAL.AdoDAL.GetInstance
            Try
                Dim dt As DataTable = oAdo.ExecuteQuery("SELECT * FROM AttributeHeader1")
            Catch ex As Exception
                'Console.WriteLine(dt.Rows.Count)
                Utilities.Utilities.LogSystemError(ex)
                Console.WriteLine(Utilities.Utilities.FilePath)
            End Try
            Console.WriteLine(oAdo.ExecuteNonQueryAsync("SELECT * FROM AttributeHeader"))
            'oAdo.ExecuteQuery("SELECT * FROM AttributeHeader", 1000)
            Console.ReadLine()
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub

End Module
