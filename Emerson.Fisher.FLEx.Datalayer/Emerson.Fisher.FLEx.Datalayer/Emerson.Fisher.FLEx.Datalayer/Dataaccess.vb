Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports Microsoft.Practices.EnterpriseLibrary.Data.Sql
Imports System.Data.Common
Public Class EmersonDAL
    Dim resultDataset As DataSet
#Region "Private Variables"
    Private m_connString As String
#End Region
#Region "common properties"
    Public Property connectionString() As String
        Get
            Return m_connString
        End Get
        Set(ByVal value As String)
            m_connString = value
        End Set
    End Property
#End Region
#Region "Stored Procedure - Data layer methods"
    ''' <summary>
    ''' Method to return the result in dataset format on execution with the specified stored procedure with the input params.
    ''' </summary>
    ''' <param name="procedureName">Procedure Name to execute</param>
    ''' <param name="inputParam">Input Parameters to the Stored Procedure</param>
    ''' <returns>returns the resultant dataset</returns>
    Public Function ExecuteDataset(ByVal procedureName As String, ByVal ParamArray inputParam As Object()) As DataSet
        Dim entLibDB As SqlDatabase = Nothing
        Dim dbCommand As DbCommand = Nothing
        Try
            entLibDB = New SqlDatabase(connectionString)
            dbCommand = entLibDB.GetStoredProcCommand(procedureName, inputParam)
            resultDataset = entLibDB.ExecuteDataSet(dbCommand)
        Catch ex As Exception
            Throw ex
        Finally
            entLibDB = Nothing
            dbCommand = Nothing
        End Try
        Return resultDataset
    End Function
    ''' <summary>
    ''' Method to return the result in datatable format on the result of execution StoredProcedure with input param
    ''' </summary>
    ''' <param name="procedureName">Procedure Name to execute</param>
    ''' <param name="inputParam">Input Parameters to the Stored Procedure</param>
    ''' <returns>returns the resultant datatable</returns>
    Public Function ExecuteDatatable(ByVal procedureName As String, ByVal ParamArray inputParam As Object()) As DataTable
        Dim resultDataTable As DataTable = Nothing
        Try
            resultDataset = ExecuteDataset(procedureName, inputParam)
            If resultDataset.Tables.Count > 0 Then
                resultDataTable = resultDataset.Tables(0)
            End If
        Catch ex As Exception
            Throw ex
        Finally
        End Try
        Return resultDataTable
    End Function
    ''' <summary>
    ''' Method to execute the Stored Procedure which will perform data manipulation operations.
    ''' </summary>
    ''' <param name="procedureName">Procedure Name to execute</param>
    ''' <param name="inputParam">Input Parameters to the Stored Procedure</param>
    ''' <returns>An indicator with the success flag of the number of affected records.</returns>
    Public Function ExecuteNonQuery(ByVal procedureName As String, ByVal ParamArray inputParam As Object()) As Integer
        Dim entLibDB As SqlDatabase = Nothing
        Dim dbCommand As DbCommand = Nothing
        Dim queryOutput As Integer
        Try
            entLibDB = New SqlDatabase(connectionString)
            dbCommand = entLibDB.GetStoredProcCommand(procedureName, inputParam)
            queryOutput = entLibDB.ExecuteNonQuery(dbCommand)
        Catch ex As Exception
            Throw ex
        Finally
            entLibDB = Nothing
            dbCommand = Nothing
        End Try
        Return queryOutput
    End Function
    ''' <summary>
    ''' Method to execute the Stored Procedure which will perform operation and return a single result.
    ''' </summary>
    ''' <param name="procedureName">Procedure Name to execute</param>
    ''' <param name="inputParam">Input Parameters to the Stored Procedure</param>
    ''' <returns>A single object which contains the results from the stored procedure execution.</returns>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(ByVal procedureName As String, ByVal ParamArray inputParam As Object()) As Object
        Dim entLibDB As SqlDatabase = Nothing
        Dim dbCommand As DbCommand = Nothing
        Dim queryOutput As Object
        Try
            entLibDB = New SqlDatabase(connectionString)
            dbCommand = entLibDB.GetStoredProcCommand(procedureName, inputParam)
            queryOutput = entLibDB.ExecuteScalar(dbCommand)

        Catch ex As Exception
            Throw ex
        Finally
            entLibDB = Nothing
            dbCommand = Nothing
        End Try
        Return queryOutput
    End Function
#End Region
#Region "SQL Inline Query - Data layer methods"
    ''' <summary>
    ''' Method to return the result in dataset format on the result of input SQL-inline Query.
    ''' </summary>
    ''' <param name="sqlQuery">Inline Query to execute</param>
    ''' <returns>returns the dataset as the ouptut of in-line query execution</returns>
    Public Function ExecuteDataset(ByVal sqlQuery As String) As DataSet
        Dim entLibDB As SqlDatabase = Nothing
        Dim dbCommand As DbCommand = Nothing
        Try
            entLibDB = New SqlDatabase(connectionString)
            dbCommand = entLibDB.GetSqlStringCommand(sqlQuery)
            resultDataset = entLibDB.ExecuteDataSet(dbCommand)
        Catch ex As Exception
            Throw ex
        Finally
            entLibDB = Nothing
            dbCommand = Nothing
        End Try
        Return resultDataset
    End Function
    ''' <summary>
    ''' Method to return the result in datatable format on the result of input SQL-inline Query.
    ''' </summary>
    ''' <param name="sqlQuery">Inline Query to execute</param>
    ''' <returns>returns the datatable as the ouptut of in-line query execution</returns>
    Public Function ExecuteDatatable(ByVal sqlQuery As String) As DataTable
        Dim resultDataTable As DataTable = Nothing
        Try
            resultDataset = ExecuteDataset(sqlQuery)
            If resultDataset.Tables.Count > 0 Then
                resultDataTable = resultDataset.Tables(0)
            End If
        Catch ex As Exception
            Throw ex
        Finally
        End Try
        Return resultDataTable
    End Function
    ''' <summary>
    ''' Method to return the datareader object on the result of input SQL-inline Query.
    ''' </summary>
    ''' <param name="sqlQuery">Inline Query to execute</param>
    ''' <returns>returns the datareader as the ouptut of in-line query execution</returns>
    Public Function ExecuteReader(ByVal sqlQuery As String) As DbDataReader
        Dim entLibDB As SqlDatabase = Nothing
        Dim dbCommand As DbCommand = Nothing
        Dim dbReader As DbDataReader = Nothing
        Try
            entLibDB = New SqlDatabase(connectionString)
            dbCommand = entLibDB.GetSqlStringCommand(sqlQuery)
            dbReader = entLibDB.ExecuteReader(dbCommand)
        Catch ex As Exception
            Throw ex
        Finally
            entLibDB = Nothing
            dbCommand = Nothing
        End Try
        Return dbReader
    End Function
    ''' <summary>
    ''' Method to return a single result on the result of input SQL-inline Query.
    ''' </summary>
    ''' <param name="sqlQuery">Inline Query to execute</param>
    ''' <returns>A single object which contains the results from the stored SQL-inline Query execution.</returns>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(ByVal sqlQuery As String) As Object
        Dim entLibDB As SqlDatabase = Nothing
        Dim dbCommand As DbCommand = Nothing
        Dim queryOutput As Object
        Try
            entLibDB = New SqlDatabase(connectionString)
            dbCommand = entLibDB.GetSqlStringCommand(sqlQuery)
            queryOutput = entLibDB.ExecuteScalar(dbCommand)
        Catch ex As Exception
            Throw ex
        Finally
            entLibDB = Nothing
            dbCommand = Nothing
        End Try
        Return queryOutput
    End Function
    ''' <summary>
    ''' Method to execute in-line query which will perform data manipulation operations.
    ''' </summary>
    '''  <param name="sqlQuery">Inline Query to execute</param>
    ''' <returns>An indicator with the success flag of the number of affected records.</returns>
    Public Function ExecuteNonQuery(ByVal sqlQuery As String) As Integer
        Dim entLibDB As SqlDatabase = Nothing
        Dim dbCommand As DbCommand = Nothing
        Dim queryOutput As Integer
        Try
            entLibDB = New SqlDatabase(connectionString)
            dbCommand = entLibDB.GetSqlStringCommand(sqlQuery)
            queryOutput = entLibDB.ExecuteNonQuery(dbCommand)
        Catch ex As Exception
            Throw ex
        Finally
            entLibDB = Nothing
            dbCommand = Nothing
        End Try
        Return queryOutput
    End Function
#End Region
End Class
