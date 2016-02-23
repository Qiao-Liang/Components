Imports System.Text
Imports System.Xml
Imports ExcelExporter.Formatter

''' <summary>
''' This class takes in data in the type of DataTable and output a string in XML spreadsheet format.
''' </summary>
''' <remarks>
''' Created by Alex Liang on 4-Apr-2011
''' V2.0 upgrade by Alex Liang on 4-Nov-2011  -- Formatter introduced, report title section enabled, DateTime bug fixed.
''' V3.0 upgrade by Alex Liang on 16-Nov-2011 -- Switched the file generating from StringBuilder to XML classes. Splited steps into multiple methods.
''' V3.1 upgrade by Alex Liang on 14-Dec-2011 -- Changes made due to the data type change of properies ColumnWidth and SearchParameters of Formatter.
''' V3.2 upgrade by Alex Liang on 11-Jan-2012 -- Introduced more number formats; Optimized the XML skeleton.
''' V4.0 upgrade by Alex Liang on 27-Jan-2012 -- Switched to the brand new template-based style settings; Built-in multi-table supported.
''' V4.1 upgrade by Alex Liang on 27-Apr-2012 -- Empty data cell removed.
''' V4.2 upgrade by Alex Liang on 3-May-2012  -- Multi-worksheet exporting enabled.
''' V4.3 upgrade by Alex Liang on 19-Jul-2012 -- Empty DataTable handled.
''' V4.4 upgrade by Alex Liang on 30-Jul-2012 -- Template property introduced to accept template in XmlDocument type; More numeric data types covered; Minor code optimization.
''' </remarks>
Public Class Exporter
    Dim xmlDoc As XmlDocument
    Dim xmlTable As XmlElement
    Dim xmlWkSht As XmlElement
    Dim xmlWkBk As XmlElement
    Dim xmlTblEnd As XmlElement
    Dim xmlNSMgr As XmlNamespaceManager
    Dim strRptNm As String
    Dim strNS As String = "urn:schemas-microsoft-com:office:spreadsheet"
    Dim iExtRow As Integer
    Dim iPage As Integer
    Dim iMaxExtRow As Integer   ' Max number of rows that can be before the row to be added

    ''' <summary>
    ''' Exports DataSet to Excel in OOXML format.
    ''' </summary>
    ''' <param name="objFmt">Instance of the ExcelFormater providing all format information</param>
    ''' <returns>The string of XML spreadsheet</returns>
    ''' <remarks>It works with XML template.</remarks>
    Public Function ExportToExcel(ByVal objFmt As Formatter) As String
        InitDoc(objFmt)
        FillFields(objFmt.Fields)
        FillData(objFmt.DataSource)

        Return xmlDoc.InnerXml
    End Function

    ''' <summary>
    ''' Initiate the XML document by loading the XML template with user defined format.
    ''' </summary>
    ''' <param name="objFmt">Instance of Formatter</param>
    ''' <remarks>Load the template and do document initialization.</remarks>
    Private Sub InitDoc(ByVal objFmt As Formatter)
        ' Initialize the XML document
        If IsNothing(objFmt.Template) Then
            xmlDoc = New XmlDocument()
            xmlDoc.Load(objFmt.TemplatePath)
        Else
            xmlDoc = objFmt.Template
        End If
        xmlNSMgr = New XmlNamespaceManager(xmlDoc.NameTable)
        xmlNSMgr.AddNamespace("ss", strNS)

        ' Initialize the mutual values
        iPage = 1
        iMaxExtRow = objFmt.MaxRowPerSheet - 1
        xmlWkBk = xmlDoc.GetElementsByTagName("Workbook")(0)   ' Only 1 workbook is included in a spreadsheet        
    End Sub

    ''' <summary>
    ''' Fill all the Fields with data.
    ''' </summary>
    ''' <param name="arlFlds">Typed value respectively matched to each Field.</param>
    ''' <remarks></remarks>
    Private Sub FillFields(ByVal arlFlds As ArrayList)
        Dim sbLocator As New StringBuilder
        Dim xmlFldLoc As XmlElement
        With sbLocator
            For iCount As Integer = 0 To arlFlds.Count - 1
                .Clear()
                .Append("/ss:Workbook/ss:Worksheet/ss:Table/ss:Row/ss:Cell/ss:Data[text()='!Field_")
                .Append(iCount)
                .Append("']")
                xmlFldLoc = xmlDoc.SelectSingleNode(.ToString, xmlNSMgr)   ' Find the locator of current Field.

                SetDataCell(xmlFldLoc, arlFlds(iCount).ToString, arlFlds(iCount).GetType.ToString)
            Next
        End With
    End Sub

    ''' <summary>
    ''' Set value and data type of data cell.
    ''' </summary>
    ''' <param name="xmlElm">The data cell to be set</param>
    ''' <param name="strVal">Value</param>
    ''' <param name="strType">Data type</param>
    ''' <remarks>2 overloads.</remarks>
    Private Sub SetDataCell(ByRef xmlElm As XmlElement, ByRef strVal As String, ByRef strType As String)
        Select Case strType
            Case "System.String"
                xmlElm.SetAttribute("Type", strNS, "String")
            Case "System.Byte", "System.SByte", "System.Int16", "System.UInt16", "System.Int32", "System.UInt32", "System.Int64", "System.UInt64", "System.Decimal", "System.Single", "System.Double"
                xmlElm.SetAttribute("Type", strNS, "Number")
            Case "System.DateTime"
                xmlElm.SetAttribute("Type", strNS, "DateTime")
                strVal = Format(CDate(strVal), "yyyy-MM-ddThh:mm:ss.sss")
        End Select

        xmlElm.InnerText = strVal
    End Sub

    ''' <summary>
    ''' Set data type of the data cell only.
    ''' </summary>
    ''' <param name="xmlElm">The data cell to be set</param>
    ''' <param name="strType">Data type</param>
    ''' <remarks>2 overloads.</remarks>
    Private Sub SetDataCell(ByRef xmlElm As XmlElement, ByRef strType As String)
        Select Case strType
            Case "System.String"
                xmlElm.SetAttribute("Type", strNS, "String")
            Case "System.Byte", "System.SByte", "System.Int16", "System.UInt16", "System.Int32", "System.UInt32", "System.Int64", "System.UInt64", "System.Decimal", "System.Single", "System.Double"
                xmlElm.SetAttribute("Type", strNS, "Number")
            Case "System.DateTime"
                xmlElm.SetAttribute("Type", strNS, "DateTime")
        End Select
    End Sub

    ''' <summary>
    ''' Fill the data.
    ''' </summary>
    ''' <param name="dsSrc">Data source in the type of DataSet</param>
    ''' <remarks>Auto-pagination is equiped.</remarks>
    Private Sub FillData(ByVal dsSrc As DataSet)
        For iCount As Integer = 0 To dsSrc.Tables.Count - 1
            AppendDataRows(GetTemplateRows(iCount), dsSrc.Tables(iCount))
        Next

        Dim xmlFstSht As XmlElement = xmlDoc.GetElementsByTagName("Worksheet")(0)
        strRptNm = xmlFstSht.Attributes("Name", strNS).Value
        If iPage > 1 Then
            Dim sbNm As New StringBuilder
            With sbNm
                .Append(strRptNm)
                .Append(" - 1")
                xmlFstSht.SetAttribute("Name", strNS, .ToString)
            End With
        End If
    End Sub

    ''' <summary>
    ''' Get the template rows of a given table.
    ''' </summary>
    ''' <param name="iLoop">Index of the table in question</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetTemplateRows(ByVal iLoop As Integer) As Array
        Dim sbLoc As New StringBuilder
        Dim arlRowTemp As New ArrayList()

        Dim xmlTblLoc As XmlElement = xmlDoc.SelectSingleNode("/ss:Workbook/ss:Worksheet/ss:Table/ss:Row/ss:Cell/ss:Data[text()='!RowTemp_Begin']", xmlNSMgr).ParentNode.ParentNode   ' Always get the locator of the first table, as the previous one is deleted after data bound.
        xmlTable = xmlTblLoc.ParentNode
        xmlTable.Attributes.RemoveAll()
        xmlWkSht = xmlTable.ParentNode
        iExtRow = GetBeforeExistRows(xmlTblLoc)
        Dim xmlRowTemp As XmlElement = xmlTblLoc.NextSibling
        xmlTable.RemoveChild(xmlTblLoc)
        While Not xmlRowTemp Is Nothing
            If xmlRowTemp.FirstChild.FirstChild.InnerText = "!RowTemp_End" Then
                xmlTblEnd = xmlRowTemp   ' Set the end row of the current table
                Exit While
            End If

            arlRowTemp.Add(xmlRowTemp)
            xmlRowTemp = xmlRowTemp.NextSibling
            xmlTable.RemoveChild(xmlRowTemp.PreviousSibling)
        End While

        Return arlRowTemp.ToArray
    End Function

    ''' <summary>
    ''' Get the existing rows before the given row.
    ''' </summary>
    ''' <param name="xmlRow">The row in question</param>
    ''' <returns>The number of exising rows before the given one</returns>
    ''' <remarks></remarks>
    Private Function GetBeforeExistRows(ByVal xmlRow As XmlElement) As Integer
        Dim iIdx As Integer = -1

        While xmlRow.Name = "Row"
            iIdx += 1
            xmlRow = xmlRow.PreviousSibling
        End While

        Return iIdx
    End Function

    ''' <summary>
    ''' Append the data rows for the current table.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AppendDataRows(ByVal arrTemp As Array, ByVal dtTbl As DataTable)
        If dtTbl.Rows.Count > 0 Then
            ' Prepare the template rows
            For Each xmlRow As XmlElement In arrTemp
                For Each xmlCell As XmlElement In xmlRow.ChildNodes
                    SetDataCell(xmlCell.FirstChild, dtTbl.Columns(xmlCell.FirstChild.InnerText).DataType.ToString)
                Next
            Next

            ' Fill the template row with data and append it to the table
            Dim xmlTemp As XmlElement
            Dim xmlData As XmlElement
            Dim iTemp As Integer = 0
            For Each drTbl As DataRow In dtTbl.Rows
                If iExtRow > iMaxExtRow Then
                    NewWorkSheet(xmlWkSht)
                End If

                xmlTemp = arrTemp(iTemp).Clone
                iTemp += 1
                If iTemp = arrTemp.Length Then
                    iTemp = 0
                End If

                For Each xmlCell As XmlElement In xmlTemp.ChildNodes
                    xmlData = xmlCell.FirstChild
                    If IsDBNull(drTbl(xmlData.InnerText)) Then
                        xmlCell.RemoveChild(xmlData)
                    Else
                        If xmlData.Attributes("Type", strNS).Value = "DateTime" Then
                            xmlData.InnerText = Format(drTbl(xmlData.InnerText), "yyyy-MM-ddThh:mm:ss.sss")
                        Else
                            xmlData.InnerText = drTbl(xmlData.InnerText).ToString
                        End If
                    End If
                Next
                xmlTable.InsertBefore(xmlTemp, xmlTblEnd)
                iExtRow += 1
            Next
        End If

        xmlTable.RemoveChild(xmlTblEnd)
    End Sub

    ''' <summary>
    ''' Create a new work sheet and append it to the work book.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub NewWorkSheet(ByVal xmlOldSht As XmlElement)
        Dim xmlCols = xmlOldSht.GetElementsByTagName("Column")
        Dim xmlOldTbl As XmlElement = xmlTable
        xmlWkSht = xmlDoc.CreateElement("Worksheet", strNS)
        xmlTable = xmlDoc.CreateElement("Table", strNS)
        ' Set the column width
        If xmlCols.Count > 0 Then
            For iCount As Integer = 0 To xmlCols.Count - 1
                xmlTable.AppendChild(xmlCols(iCount).Clone)
            Next
        End If

        Dim xmlTail As XmlElement = xmlTblEnd
        While Not xmlTail Is Nothing
            xmlOldTbl.RemoveChild(xmlTail)
            xmlTable.AppendChild(xmlTail.Clone)
            xmlTail = xmlTail.NextSibling
        End While
        xmlTblEnd = xmlTable.GetElementsByTagName("Row")(0)   ' Specifies the table end indicator

        iPage += 1
        Dim sbName As New StringBuilder
        With sbName
            .Append(strRptNm)
            .Append(" - ")
            .Append(iPage)

            xmlWkSht.SetAttribute("Name", strNS, .ToString)
        End With
        iExtRow = 0
        xmlWkSht.AppendChild(xmlTable)
        xmlWkBk.AppendChild(xmlWkSht)
    End Sub
End Class