Public Class Tester
    Public Sub Main()
        Dim objExl As New ExcelExporter
        Dim dtTest As New DataTable
        dtTest.Columns.Add("Header1")
        dtTest.Columns.Add("Header2")
        dtTest.Columns.Add("Header3")
        Dim drTest As DataRow = dtTest.NewRow
        drTest(0) = 0
        drTest(1) = 1
        drTest(2) = 2
        dtTest.Rows.Add(drTest)
        Dim arrTest(3) As Integer
        arrTest(0) = 0
        arrTest(1) = 1
        arrTest(2) = 2

        objExl.ExportToExcel(dtTest, "Test", arrTest)
    End Sub
End Class
