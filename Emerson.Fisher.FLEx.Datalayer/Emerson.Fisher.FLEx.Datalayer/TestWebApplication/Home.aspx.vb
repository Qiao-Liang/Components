Imports Emerson.Fisher.Flex.Datalayer

Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    'Protected Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
    '    Dim sqlquery As String = "select * from InventoryOrgs"
    '    Dim objDBDAL As New EmersonDAL()
    '    Dim outputDS As DataSet = Nothing
    '    objDBDAL.connectionString = GetConnectionString()
    '    outputDS = objDBDAL.ExecuteDataset(sqlquery)
    '    Dim outputTables As Integer = outputDS.Tables.Count


    'End Sub
    Private Function GetConnectionString() As String
        Return ConfigurationManager.ConnectionStrings("ConnString").ConnectionString.ToString
    End Function

    Protected Sub btnCauseExc_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCauseExc.Click
        Dim x, y, z As Integer
        y = 1
        z = 0
        x = y / z
    End Sub
End Class