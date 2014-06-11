Public Class AdoConstants
    Public Const MaxLargeDataLength As Int64 = 500000000

    Public Enum DatabaseTypes
        [Nothing] = -1
        Oracle = 0
        MSSql = 1
    End Enum
    Public Enum DataTypes
        [Char] = 0
        [VarChar] = 1
        [Integer] = 2
        [Double] = 3
        [DateTime] = 4
        [LargeText] = 5
        [Image] = 6
    End Enum

End Class

