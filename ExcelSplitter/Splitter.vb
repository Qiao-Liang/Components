Imports System.Text
Imports System.Xml
Imports ExcelExporter

''' <summary>
''' A plug-in for the ExcelExport component.
''' </summary>
''' <remarks>Created by Alex Liang on 9-Nov-2011. As the first version, it doesn't cover the pagination.</remarks>
Public Class Splitter
    ''' <summary>
    ''' Export data splitted into multiple tables to Excel spreadsheet.
    ''' </summary>
    ''' <param name="dtSplt">Data source</param>
    ''' <param name="objFmt">Instance of Formater</param>
    ''' <param name="arrGroup">Array of columns by which the data rows are grouped</param>
    ''' <returns>String in XML manner</returns>
    ''' <remarks>The input DataTable would be sorted by the given columns.</remarks>
    Public Function ExportToSplitExcel(ByVal dtSplt As DataTable, ByVal objFmt As Formater, ByVal arrGroup As String()) As String
        ' Sort the input DataTable
        Dim sbSort As New StringBuilder
        Dim objColHash As New Hashtable
        Dim iArrLth As Integer = arrGroup.Length
        For iCount As Integer = 0 To iArrLth - 1
            sbSort.Append(",")
            sbSort.Append(arrGroup(iCount))
        Next
        sbSort.Remove(0, 1)   ' Remove the first coma
        Dim dvSort As New DataView(dtSplt)
        dvSort.Sort = sbSort.ToString
        Dim dtSort As DataTable = dvSort.ToTable

        ' Get the index of the first rows of each groups
        Dim iIdx As Integer = 0
        Dim arlStrIdx As New ArrayList   ' Hold the start index
        For Each drSort As DataRow In dtSort.Rows
            If iIdx > 0 Then
                For iCount As Integer = 0 To iArrLth - 1
                    If drSort(arrGroup(iCount)) <> dtSort.Rows(iIdx - 1)(arrGroup(iCount)) Then
                        arlStrIdx.Add(iIdx)
                        Exit For
                    End If
                Next
            End If
            iIdx += 1
        Next

        ' Get the original OOXML spreadsheet by ExcelExporter.Exporter
        objFmt.ShowHeader = False   ' Force to hide the header
        Dim objExp As New Exporter
        Dim strXML = objExp.ExportToExcel(dtSort, objFmt)

        Dim xmlDoc As New XmlDocument
        Dim strNS As String = "urn:schemas-microsoft-com:office:spreadsheet"
        xmlDoc.LoadXml(strXML)
        Dim iOffSet As Integer = 0   ' Grab the number of rows occupied by the title section (report name and parameters)
        If Not objFmt.SearchParameters Is Nothing Then
            iOffSet = objFmt.SearchParameters.Count + 1
        End If

        ' Get the hash mapping of the DataTable row index and OOXML rows.
        arlStrIdx.Insert(0, 0)   ' Add the 0 index
        Dim xmlRows As XmlNodeList = xmlDoc.GetElementsByTagName("Row")
        Dim hshStrRows As New Hashtable
        For Each iRowIdx As Integer In arlStrIdx.ToArray(GetType(System.Int32))
            hshStrRows.Add(iRowIdx, xmlRows(iRowIdx + iOffSet))
        Next

        ' Prepare the row templates
        Dim xmlHeader As XmlElement = xmlDoc.CreateElement("Row", strNS)
        For Each dcSort As DataColumn In dtSort.Columns
            Dim xmlCell As XmlElement = xmlDoc.CreateElement("Cell", strNS)
            Dim xmlData As XmlElement = xmlDoc.CreateElement("Data", strNS)
            xmlCell.SetAttribute("StyleID", strNS, "sHeader")
            xmlData.SetAttribute("Type", strNS, "String")
            xmlData.InnerText = dcSort.ColumnName
            xmlCell.AppendChild(xmlData)
            xmlHeader.AppendChild(xmlCell)
        Next
        Dim xmlTable As XmlElement = xmlDoc.GetElementsByTagName("Table")(0)
        Dim xmlEmptyRow As XmlElement = xmlDoc.CreateElement("Row", strNS)
        Dim xmlEmptyCell As XmlElement = xmlDoc.CreateElement("Cell", strNS)
        Dim xmlMergeRow As XmlElement = xmlDoc.CreateElement("Row", strNS)
        Dim xmlMergeCell As XmlElement = xmlDoc.CreateElement("Cell", strNS)
        Dim xmlMergeData As XmlElement = xmlDoc.CreateElement("Data", strNS)
        xmlEmptyCell.SetAttribute("MergeAcross", strNS, dtSort.Columns.Count - 1)
        xmlMergeCell.SetAttribute("MergeAcross", strNS, dtSort.Columns.Count - 1)
        xmlMergeData.SetAttribute("Type", strNS, "String")
        xmlMergeCell.AppendChild(xmlMergeData)
        xmlMergeRow.AppendChild(xmlMergeCell)
        xmlEmptyRow.AppendChild(xmlEmptyCell)

        ' Insert breaking row (an empty row), table title and header to each of the groups splitted from the DataTable.
        Dim xmlClnMergeRow As XmlElement
        Dim sbRowTitle As New StringBuilder
        Dim drKey As DataRow
        Dim strColNm As String
        For Each iKey As Integer In hshStrRows.Keys
            sbRowTitle.Clear()
            drKey = dtSort.Rows(iKey)
            For iCount As Integer = 0 To iArrLth - 1
                strColNm = arrGroup(iCount)
                sbRowTitle.Append(" / ")
                sbRowTitle.Append(strColNm)
                sbRowTitle.Append(" - ")
                sbRowTitle.Append(drKey(strColNm))
            Next
            sbRowTitle.Remove(0, 3)   ' Remove the leading slash.

            xmlClnMergeRow = xmlMergeRow.Clone
            xmlClnMergeRow.ChildNodes(0).ChildNodes(0).InnerText = sbRowTitle.ToString

            Dim xmlBefore As XmlNode = hshStrRows(iKey)
            xmlTable.InsertBefore(xmlEmptyRow.Clone, xmlBefore)
            xmlTable.InsertBefore(xmlClnMergeRow, xmlBefore)
            xmlTable.InsertBefore(xmlHeader.Clone, xmlBefore)
        Next

        ' Reset the number of total rows
        xmlTable.SetAttribute("ExpandedRowCount", strNS, xmlTable.GetElementsByTagName("Row").Count)

        Return xmlDoc.InnerXml
    End Function
End Class
