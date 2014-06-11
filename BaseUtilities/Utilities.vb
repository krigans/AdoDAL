Public Class Utilities
    Public Shared Function Equal(ByVal oObject1 As Object, ByVal oObject2 As Object, Optional ByVal bTreatZeroAsNothing As Boolean = False) As Boolean
        Dim bIsEqual As Boolean = False
        Try
            If oObject1 Is Nothing OrElse TypeOf (oObject1) Is DBNull OrElse (TypeOf (oObject1) Is String And oObject1.ToString().Trim() = String.Empty) Then
                If oObject2 Is Nothing OrElse TypeOf (oObject2) Is DBNull OrElse (TypeOf (oObject2) Is String And oObject2.ToString().Trim() = String.Empty) Then bIsEqual = True Else bIsEqual = False
            Else
                If oObject2 Is Nothing OrElse TypeOf (oObject2) Is DBNull OrElse (TypeOf (oObject2) Is String And oObject2.ToString().Trim() = String.Empty) Then
                    bIsEqual = False
                Else
                    If TypeOf (oObject1) Is DateTime Then
                        bIsEqual = IIf(CType(CType(oObject1, DateTime).Subtract(CType(oObject2, DateTime)), TimeSpan).TotalMilliseconds = 0, True, False)
                    ElseIf TypeOf (oObject1) Is String Then
                        'REASONING: Since Database Char DataType pads the actual string value with spaces to the right, we would have to trim extra spaces before equating.
                        bIsEqual = oObject1.ToString().Trim().Equals(oObject2.ToString().Trim())
                    Else
                        'REASONING: Done so to equate values instead of references
                        'ASSUMPTION: ToString() would act in the same manner in both objects as Type is same or similar
                        'ASSUMPTION: ToString() would convert the Value of the object to string format without loss of precision and hence check for Value equality.
                        bIsEqual = oObject1.ToString().Equals(oObject2.ToString())
                    End If
                End If
            End If
            If Not bIsEqual AndAlso Not oObject1 Is Nothing Then
                Select Case oObject1.GetType.ToString
                    Case "System.DateTime"
                        If oObject1 = DateTime.MinValue AndAlso (oObject2 Is Nothing OrElse TypeOf (oObject2) Is DBNull) Then
                            bIsEqual = True
                        End If
                    Case "System.Int16", "System.Int32", "System.Int64", "System.Integer", "System.Single", "System.Double"
                        If bTreatZeroAsNothing And oObject1 = 0 AndAlso (oObject2 Is Nothing OrElse TypeOf (oObject2) Is DBNull) Then
                            bIsEqual = True
                        End If
                End Select
            End If
            If Not bIsEqual AndAlso Not oObject2 Is Nothing Then
                Select Case oObject2.GetType.ToString
                    Case "System.DateTime"
                        If oObject2 = DateTime.MinValue AndAlso (oObject1 Is Nothing OrElse TypeOf (oObject1) Is DBNull) Then
                            bIsEqual = True
                        End If
                    Case "System.Int16", "System.Int32", "System.Int64", "System.Integer", "System.Single", "System.Double"
                        If bTreatZeroAsNothing And oObject2 = 0 AndAlso (oObject1 Is Nothing OrElse TypeOf (oObject1) Is DBNull) Then
                            bIsEqual = True
                        End If
                End Select
            End If
        Catch ex As Exception
            'LogSystemError(ex)
        End Try
        Return bIsEqual
    End Function
    Public Shared Function Decrypt(ByVal sEncryptedPassword As String) As String
        Dim sDecryptedPassword As String = Nothing
        Try
            If Not Equal(sEncryptedPassword, Nothing) Then
                Dim aEncryptedPasswordStrings() As String
                aEncryptedPasswordStrings = sEncryptedPassword.Split("~")
                Dim aEncryptedPassword(aEncryptedPasswordStrings.Length - 1) As Byte
                For iCharIndex As Integer = 0 To (aEncryptedPasswordStrings.Length - 1)
                    aEncryptedPassword(iCharIndex) = Convert.ToByte(aEncryptedPasswordStrings(iCharIndex))
                Next
                Dim oRSACryptoServiceProvider As System.Security.Cryptography.RSACryptoServiceProvider = New System.Security.Cryptography.RSACryptoServiceProvider
                oRSACryptoServiceProvider.FromXmlString("<RSAKeyValue><Modulus>wqC0CH/Jk141//A2f3BKU8QHX9cy+iqTVHlBm4jsaiux6SA9+yX6v1la3c2qhqecdzQ8rR5xTg9ro2lJZ0l8CLTXGC4Llz3LfcWx8SLFuc2yM53ICCOaO4jB/vJkW7CzNKQGwdLBSiYv1/45H6blQIK9/zBGVCeDqw5H9WCmZ1E=</Modulus><Exponent>AQAB</Exponent><P>9FMbInveHHDW0ds/Dpv6crVwVWTWiS4w3cm7vzBh3+vRZux0yImAa0U/H4pelDEABrTm9QfKSlFE3ra+3ovgSw==</P><Q>y+2kD4A8SicCjU4VywZOJvTCjIE3l54AoDAqmCGft8g6HF0MLv01axq5qPOWl4KPrzvzh2BYykfJnNdyZ+6tUw==</Q><DP>whMJqNiv0/OmEEiRzC8GP/vz4UEaURmJ44MNSY9LD62oRpNpKKpggdUdkRY+joRluu4Tz2uCuonXpPmQoAKIBQ==</DP><DQ>ovahKJH9m/RYobtIxxme0pq97bJFTrBBJ8HWCAS2shMb/RaOae6HBbQxscYXDbSURiDOl9xymBOOFfxFvLCLaQ==</DQ><InverseQ>1Lc5vJUlyAPWwXtwHGPygRYphD/NthlAqnywNA4t8Pg6XTWV4whnLr1lp2C71AnyWGDkglmEBvIIkVNvUDXzRw==</InverseQ><D>RMg4d9x5Z5xe5yGEkQslKW9Yz9UkzeZoBO2Jcyczrd3dVS8w2GY2tJMmmsaJYmcv06zhWKkuj9DBUJHwABGnRGoCDFagPpVrScl2tAzg+qvRvDiqBwoguMnQp2l+VdsJqgZI+/e+PxXpt8+sYroA+KOYSEVM2/j8cvMjtZsuLhU=</D></RSAKeyValue>")
                Dim aDecryptedPassword() As Byte = oRSACryptoServiceProvider.Decrypt(aEncryptedPassword, False)
                For iByteIndex As Integer = 0 To aDecryptedPassword.Length - 1
                    sDecryptedPassword = sDecryptedPassword + Convert.ToChar(aDecryptedPassword(iByteIndex))
                Next
            End If
        Catch ex As Exception
            ''LogSystemError(ex)
            Throw ex
        End Try
        Console.WriteLine(sDecryptedPassword)
        Return sDecryptedPassword
    End Function
    Public Shared Sub LogSystemError(ex As Exception)
        Console.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
    End Sub

    Public Shared Sub LogSystemError(ex As Exception, sErrorMessage As String)
        Console.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
    End Sub

End Class
