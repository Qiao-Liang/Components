Imports Emerson.Fisher.Flex.Datalayer
Imports System.Data.SqlClient

Public Class SP_QueryTesting
    Inherits System.Web.UI.Page
    Private objDBDAL As New EmersonDAL()
    Private Function GetConnectionString(ByVal connstr As String) As String
        Return ConfigurationManager.ConnectionStrings(connstr).ConnectionString.ToString
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            LoadServiceReqInfo()
        End If
    End Sub
    Private Sub GetLineNotes()
        Dim sqlProcName As String = "FLEX.SPOM_LINENOTES"
        Dim lineId As String = txtLineId.Text
        Dim lbpInfo As String = txtInternalInd.Text
        Dim outputResult As Object = Nothing
        objDBDAL.connectionString = GetConnectionString("OrderManagement")
        Try
            outputResult = objDBDAL.ExecuteDatatable(sqlProcName, GetLineId(), GetIO5Ind(), GetInternalInd())
            lblLineNotes.Text = "Lines notes results : "
            lgSPTesting.InnerText = lgSPTesting.InnerText + " Procedure Name : " + sqlProcName
            grdViewLineNotes.DataSource = outputResult
            grdViewLineNotes.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadServiceReqInfo()
        Dim sqlProcName As String = "FLEX.SP_InsertSelectServiceReq"
        Dim outputResult As Object = Nothing
        objDBDAL.connectionString = GetConnectionString("ServiceRequestConnectionString")
        Try
            outputResult = objDBDAL.ExecuteDatatable(sqlProcName, "S", GetDBNullValue(), GetDBNullValue(), GetDBNullValue(), GetDBNullValue(), GetDBNullValue())
            grdProbDat.DataSource = outputResult
            grdProbDat.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function GetDBNullValue() As Object
        Return System.DBNull.Value
    End Function
    Private Function GetLineId() As Object
        If Not txtLineId.Text = String.Empty Then
            Return Convert.ToInt32(txtLineId.Text)
        Else
            Return GetDBNullValue()
        End If
    End Function
    Private Function GetInternalInd() As Object
        If Not txtInternalInd.Text = String.Empty Then
            Return Convert.ToInt32(txtInternalInd.Text)
        Else
            Return GetDBNullValue()
        End If
    End Function
    Private Function GetIO5Ind() As Object
        If Not txtIO5Indicator.Text = String.Empty Then
            Return Convert.ToInt32(txtIO5Indicator.Text)
        Else
            Return GetDBNullValue()
        End If
    End Function

    ''' <summary>
    ''' for testing sp when the there is none parament.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub TestSP()
        Dim sqlsp As String = "TestGetLBP"
        Dim inputParam() As Object = Nothing
        Dim outputDS As DataSet = Nothing
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        outputDS = objDBDAL.ExecuteDataset(sqlsp, inputParam)
        If outputDS.Tables.Count > 0 Then
            GridView1.DataSource = outputDS
            GridView1.DataBind()
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Dim sqlsp As String = "FLEX.SPSR_GetLBPName"
        Dim outputDS As DataSet = Nothing
        objDBDAL.connectionString = GetConnectionString("ServiceRequestConnectionString")
        outputDS = objDBDAL.ExecuteDataset(sqlsp, "013")
        If outputDS.Tables.Count > 0 Then
            GridView1.DataSource = outputDS
            GridView1.DataBind()
        End If

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        GetMutipleTables()
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
    End Sub

    Protected Sub GetMutipleTables()
        Dim sqlsp As String = "MutipleReturn"
        Dim outputDS As DataSet = Nothing
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        outputDS = objDBDAL.ExecuteDataset(sqlsp, "0")
        If outputDS.Tables.Count > 1 Then
            GridView2.DataSource = outputDS.Tables(0)
            GridView2.DataBind()
            DropDownList1.Visible = True
            DropDownList1.DataSource = outputDS.Tables(1)
            DropDownList1.DataTextField = "LBPID"
            DropDownList1.DataValueField = "LBPID"
            DropDownList1.DataBind()

        ElseIf outputDS.Tables.Count > 0 Then
            GridView2.DataSource = outputDS.Tables(0)
            GridView2.DataBind()
        End If
    End Sub
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        Dim sqlsp As String = "InsertDATA"
        Dim outputDS As DataSet = Nothing
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        Try
            objDBDAL.ExecuteNonQuery(sqlsp, "0")
            Label1.Text = "Insert status :insert success."
            GetMutipleTables()
        Catch ex As Exception
            Label1.Text = "Insert status :insert failure."
        End Try

    End Sub

    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button5.Click
        Dim sqlsp As String = "updatedate"
        Dim outputDS As DataSet = Nothing
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        Try
            Dim lbp As Integer = Int(DropDownList1.SelectedValue)
            objDBDAL.ExecuteNonQuery(sqlsp, lbp)
            Label3.Text = "Update statu:update success."
            GetMutipleTables()
        Catch ex As Exception
            Label3.Text = "Update statu:update unsuccess."
        End Try
    End Sub

    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button6.Click
        Dim sqlsp As String = "deletedate"
        Dim outputDS As DataSet = Nothing
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        Try
            Dim lbp As Integer = Int(DropDownList1.SelectedValue)
            objDBDAL.ExecuteNonQuery(sqlsp, lbp)
            Label4.Text = "Delete statu:delete success."
            GetMutipleTables()
        Catch ex As Exception
            Label4.Text = "Delete statu:delete unsuccess."
        End Try
    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button4.Click
        Dim sqlsp As String = "InserandReturnPrimay"
        Dim outputDS As DataSet = Nothing
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        Try
            outputDS = objDBDAL.ExecuteDataset(sqlsp, "0")
            If outputDS.Tables.Count > 0 Then
                Label2.Text = "Inser result:insert success and return " & outputDS.Tables(0).Rows(0)(0) & "."
            End If
            GetMutipleTables()
        Catch ex As Exception
            Label2.Text = "Insert result:update failure."
        End Try
    End Sub

    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button7.Click
        Dim sqlsp As String = "InserandReturnPrimay"
        Dim outputDS As Object
        objDBDAL.connectionString = GetConnectionString("localconnstring")
        Try
            outputDS = objDBDAL.ExecuteScalar(sqlsp, "0")
            Label5.Text = "test ExecuteScalar: return " & outputDS & "."
            GetMutipleTables()
        Catch ex As Exception
            Label5.Text = "test ExecuteScalar: failure."
        End Try
    End Sub

    Protected Sub btnLineNotes_click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLineNotes.Click
        GetLineNotes()
    End Sub

    Protected Sub btnProbInsert_click(ByVal sender As Object, ByVal e As EventArgs) Handles btnProbInsert.Click
        InsertServiceRequestData()
    End Sub
    Private Sub InsertServiceRequestData()
        Dim sqlProcName As String = "FLEX.SP_InsertSelectServiceReq"
        Dim outputResult As Object = Nothing
        objDBDAL.connectionString = GetConnectionString("ServiceRequestConnectionString")
        Try
            outputResult = objDBDAL.ExecuteDatatable(sqlProcName, "I", "Infosys", GetProdType(), GetProdQuantity(), GetProdDescription(), GetProbType())
            grdProbDat.DataSource = outputResult
            grdProbDat.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function GetProdType() As Object
        If Not txtProdType.Text = String.Empty Then
            Return Convert.ToString(txtProdType.Text)
        Else
            Return GetDBNullValue()
        End If
    End Function
    Private Function GetProdQuantity() As Object
        If Not txtQuantity.Text = String.Empty Then
            Return Convert.ToInt32(txtQuantity.Text)
        Else
            Return GetDBNullValue()
        End If
    End Function
    Private Function GetProdDescription() As Object
        If Not txtProbDesc.Text = String.Empty Then
            Return Convert.ToString(txtProbDesc.Text)
        Else
            Return GetDBNullValue()
        End If
    End Function
    Private Function GetProbType() As Object
        If Not txtProbType.Text = String.Empty Then
            Return Convert.ToString(txtProbType.Text)
        Else
            Return GetDBNullValue()
        End If
    End Function
End Class