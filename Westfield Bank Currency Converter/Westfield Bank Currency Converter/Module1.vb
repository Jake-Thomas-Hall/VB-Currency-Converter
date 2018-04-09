Module Module1
    'Declares a variable with 8 items for the module
    Public PercentageRatesArray(0 To 8) As Decimal
    'Declares a name for the module
    Public Sub LoadRates()
        'StreamReader is used to read the textfile
        Dim Rates As New System.IO.StreamReader("C:\Users\jake-\Documents\L3 Dip Applied Computing\Unit 614 VB Programs\Westfield Bank Currency Converter\PercentageRates.txt")
        'For statement iterates 8 times, reading each line and then inserting this into the array.
        For i As Integer = 0 To 8
            PercentageRatesArray(i) = Rates.ReadLine()
        Next
        'Object file is then closed
        Rates.Close()
    End Sub
End Module
