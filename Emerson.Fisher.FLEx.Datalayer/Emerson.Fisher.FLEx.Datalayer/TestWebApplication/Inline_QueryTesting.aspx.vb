Imports Emerson.Fisher.Flex.Datalayer
Imports System.Data.Common

Public Class Inline_QueryTesting
    Inherits System.Web.UI.Page
    Private objDBDAL As New EmersonDAL()
    Dim _hasProcessedHeader As Boolean
    Private Function GetConnectionString(ByVal connstr As String) As String
        Return ConfigurationManager.ConnectionStrings(connstr).ConnectionString.ToString
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            'RefreshLBP()
            'GetlbpInfo()
            LoadInventoryOrgs()
            LoadOrderInfo_Local()
        End If
    End Sub
    Private Sub LoadInventoryOrgs()
        Dim sqlquery As String = " SELECT ORGANIZATION_CODE,ORGANIZATION_ID,OPERATING_UNIT,FUNCTIONINGCURRENCYCODE,STANDARD_GMT_DEVIATION_HOURS,DAYLIGHT_SAVINGS_TIME_FLAG FROM INVENTORYORGS"
        Dim outputDt As DataTable = Nothing
        objDBDAL.connectionString = GetConnectionString("ConnString")
        Try
            outputDt = objDBDAL.ExecuteDatatable(sqlquery)
            If Not outputDt Is Nothing Then
                lgInvOrg.InnerText = lgInvOrg.InnerText + " : " + sqlquery
                lvInventoryOrgs.DataSource = outputDt
                lvInventoryOrgs.DataBind()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub btnGetPO_click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGetPO.Click
        Dim orderNumber As Integer = Convert.ToInt32(txtOrderNumber.Text)

        Dim sqlquery As String = " SELECT CUST_PO_NUMBER FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = " & orderNumber
        Dim dispValue As String = "Customer PO Number is "
        Dim custPoNumber As Object = Nothing
        objDBDAL.connectionString = GetConnectionString("ConnString")
        Try
            custPoNumber = objDBDAL.ExecuteScalar(sqlquery)
            lgExecuteReader.InnerText = lgExecuteReader.InnerText + " : " + sqlquery
            dispValue = dispValue + Convert.ToString(custPoNumber)
            lblPONumber.Text = dispValue
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub btnDeleteOrder_click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDeleteOrder.Click
        Dim orderNumber As Integer = Convert.ToInt32(txtOrderNoToDelete.Text)
        Dim sqlquery As String = " DELETE FROM LOCAL_OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = " & orderNumber
        Dim dispValue As String = "Delete Transaction : "
        Dim tranOutput As Integer
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        Try
            tranOutput = objDBDAL.ExecuteNonQuery(sqlquery)
            lgNonQuery.InnerText = lgNonQuery.InnerText + " : " + sqlquery
            If tranOutput > 0 Then
                lblDeleteStatus.Text = txtOrderNoToDelete.Text.Trim() + " Order information deleted successfully."
                LoadOrderInfo_Local()
            Else
                lblDeleteStatus.Text = txtOrderNoToDelete.Text.Trim() + " Order information not exists in database. Please contact admin"
            End If
        Catch ex As Exception
            lblDeleteStatus.Text = "Error in deleting " + txtOrderNoToDelete.Text.Trim() + " Order information. Please contact admin"
        End Try
    End Sub
    Private Sub LoadOrderInfo_Local()
        Dim sqlquery As String = "SELECT ORDER_NUMBER,CUST_PO_NUMBER,SHIPPING_METHOD_CODE,PACKING_INSTRUCTIONS FROM LOCAL_OE_ORDER_HEADERS_ALL"
        Dim outputDt As DataTable = Nothing
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        Try
            outputDt = objDBDAL.ExecuteDatatable(sqlquery)
            If Not outputDt Is Nothing Then
                lvLocalOrderInfo.DataSource = outputDt
                lvLocalOrderInfo.DataBind()
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class