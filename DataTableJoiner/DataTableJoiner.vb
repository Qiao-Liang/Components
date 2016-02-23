''' <summary>
''' This class helps with DataTable joining.
''' </summary>
''' <remarks>
''' Version 1.0: Alex Liang 20-Jul-2011
''' </remarks>
Public Class DataTableJoiner
    Dim dsJoin As DataSet
    Dim drlJoin As DataRelation
    Dim dtJoin As DataTable

    ''' <summary>
    ''' This method does table Join.
    ''' </summary>
    ''' <param name="dtL">Left table</param>
    ''' <param name="dtR">Right table</param>
    ''' <param name="dcL">Referrence column of the left table</param>
    ''' <param name="dcR">Referrence column of the right table</param>
    ''' <param name="bCons">A boolean value to specify whether data constraint to be set up</param>
    ''' <returns>A DataTable instance with joined data</returns>
    ''' <remarks></remarks>
    Public Function Join(ByVal dtL As DataTable, ByVal dtR As DataTable, ByVal dcL As DataColumn, ByVal dcR As DataColumn, ByVal bCons As Boolean) As DataTable
        dtJoin = JoinTableStrc(dtL, dtR, dcL, dcR, "Join", bCons)

        For Each dr As DataRow In drlJoin.ChildTable.Rows
            Dim drJoin As DataRow = dtJoin.NewRow
            Dim drPrnt As DataRow = dr.GetParentRow("Join")

            For Each dc As DataColumn In drPrnt.Table.Columns
                drJoin(dc.ColumnName) = drPrnt(dc.ColumnName)
            Next
            For Each dc As DataColumn In dr.Table.Columns
                drJoin(dc.ColumnName) = dr(dc.ColumnName)
            Next

            dtJoin.Rows.Add(drJoin)
        Next

        ReleaseTable("Join", drlJoin, bCons)

        Return dtJoin
    End Function

    ''' <summary>
    ''' This method does left outer join.
    ''' </summary>
    ''' <param name="dtL">Left table</param>
    ''' <param name="dtR">Right table</param>
    ''' <param name="dcL">Referrence column of the left table</param>
    ''' <param name="dcR">Referrence column of the right table</param>
    ''' <param name="bCons">A boolean value to specify whether data constraint to be set up</param>
    ''' <returns>A DataTable instance with joined data</returns>
    ''' <remarks></remarks>
    Public Function LeftOuterJoin(ByVal dtL As DataTable, ByVal dtR As DataTable, ByVal dcL As DataColumn, ByVal dcR As DataColumn, ByVal bCons As Boolean) As DataTable
        dtJoin = JoinTableStrc(dtL, dtR, dcL, dcR, "LeftOuterJoin", bCons)

        DoOuterJoin("LeftOuterJoin")

        ReleaseTable("LeftOuterJoin", drlJoin, bCons)

        Return dtJoin
    End Function

    ''' <summary>
    ''' This method does right outer join.
    ''' </summary>
    ''' <param name="dtL">Left table</param>
    ''' <param name="dtR">Right table</param>
    ''' <param name="dcL">Referrence column of the left table</param>
    ''' <param name="dcR">Referrence column of the right table</param>
    ''' <param name="bCons">A boolean value to specify whether data constraint to be set up</param>
    ''' <returns>A DataTable instance with joined data</returns>
    ''' <remarks></remarks>
    Public Function RightOuterJoin(ByVal dtL As DataTable, ByVal dtR As DataTable, ByVal dcL As DataColumn, ByVal dcR As DataColumn, ByVal bCons As Boolean) As DataTable
        dtJoin = JoinTableStrc(dtL, dtR, dcL, dcR, "RightOuterJoin", bCons)

        DoOuterJoin("RightOuterJoin")

        ReleaseTable("RightOuterJoin", drlJoin, bCons)

        Return dtJoin
    End Function

    ''' <summary>
    ''' The logic of outer joinings.
    ''' </summary>
    ''' <param name="strRelName">Name of the data relation</param>
    ''' <remarks>It works for both left outer join and right outer join.</remarks>
    Private Sub DoOuterJoin(ByVal strRelName As String)
        For Each dr As DataRow In drlJoin.ParentTable.Rows
            Dim drJoin As DataRow = dtJoin.NewRow
            Dim drChd() As DataRow = dr.GetChildRows(strRelName)

            For Each dc As DataColumn In dr.Table.Columns
                drJoin(dc.ColumnName) = dr(dc.ColumnName)
            Next
            If drChd.Length = 1 Then
                For Each dc As DataColumn In drChd(0).Table.Columns
                    drJoin(dc.ColumnName) = drChd(0)(dc.ColumnName)
                Next
            End If

            dtJoin.Rows.Add(drJoin)
        Next
    End Sub

    ''' <summary>
    ''' This method constructs the structure of the joined table.
    ''' </summary>
    ''' <param name="dtL">Left table</param>
    ''' <param name="dtR">Right table</param>
    ''' <param name="dcL">Referrence column of the left table</param>
    ''' <param name="dcR">Referrence column of the right table</param>
    ''' <param name="strRelName">Name of the data relation</param>
    ''' <param name="bCons">A boolean value to specify whether data constraint to be set up</param>
    ''' <returns>An empty DataTable instance covering all the columns of both the left and right table.</returns>
    ''' <remarks></remarks>
    Private Function JoinTableStrc(ByVal dtL As DataTable, ByVal dtR As DataTable, ByVal dcL As DataColumn, ByVal dcR As DataColumn, ByVal strRelName As String, ByVal bCons As Boolean) As DataTable
        dsJoin = New DataSet

        dsJoin.Tables.Add(dtL)
        dsJoin.Tables.Add(dtR)

        Dim dtJoin As New DataTable

        For Each dc As DataColumn In dtL.Columns
            dtJoin.Columns.Add(dc.ColumnName, dc.DataType)
        Next
        For Each dc As DataColumn In dtR.Columns
            dtJoin.Columns.Add(dc.ColumnName, dc.DataType)
        Next

        Select Case strRelName
            Case "Join"
                If dtL.Rows.Count >= dtR.Rows.Count Then
                    drlJoin = New DataRelation(strRelName, dcL, dcR, bCons)
                Else
                    drlJoin = New DataRelation(strRelName, dcR, dcL, bCons)
                End If
            Case "LeftOuterJoin"
                drlJoin = New DataRelation(strRelName, dcL, dcR, bCons)
            Case "RightOuterJoin"
                drlJoin = New DataRelation(strRelName, dcR, dcL, bCons)
        End Select

        dsJoin.Relations.Add(drlJoin)

        Return dtJoin
    End Function

    ''' <summary>
    ''' Release table from the DataSet with DataRelation setup and dispose this Data Set.
    ''' </summary>
    ''' <param name="strRel">Name of the data relation</param>
    ''' <param name="drlJoin">Instance of DataRelation set up in the DataSet</param>
    ''' <param name="bCons">A boolean value to specify whether data constraint is set up</param>
    ''' <remarks>The data relation is removed at first, then the data constraints on tables if exists, at last, the DataSet instance is disposed.</remarks>
    Private Sub ReleaseTable(ByVal strRel As String, ByVal drlJoin As DataRelation, ByVal bCons As Boolean)
        Dim strChdTbl As String = drlJoin.ChildTable.TableName
        Dim strPrntTbl As String = drlJoin.ParentTable.TableName

        dsJoin.Relations.Remove(strRel)

        If bCons Then
            dsJoin.Tables(strChdTbl).Constraints.RemoveAt(0)
            dsJoin.Tables(strPrntTbl).Constraints.RemoveAt(0)
        End If

        dsJoin.Tables.Remove(strChdTbl)
        dsJoin.Tables.Remove(strPrntTbl)

        dsJoin.Dispose()
    End Sub
End Class
