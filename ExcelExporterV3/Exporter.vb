Imports System.Text
Imports System.Reflection
Imports System.Xml
Imports ExcelExporter.Formater

''' <summary>
''' This class takes in data in the type of DataTable and output a string in XML spreadsheet format.
''' </summary>
''' <remarks>
''' Created by Alex Liang on 4-Apr-2011
''' V2.0 upgrade by Alex Liang on 4-Nov-2011 -- Exporter introduced, report title section enabled,  DateTime bug fixed.
''' V3.0 upgrade by Alex Liang on 16-Nov-2011 -- Switched the file generating from StringBuilder to XML classes. Splited steps into multiple methods.
''' V3.1 upgrade by Alex Liang on 14-Dec-2011 -- Changes made due to the data type change of properies ColumnWidth and SearchParameters of Formater.
''' V3.2 upgrade by Alex Liang on 11-Jan-2012 -- Introduced more number formats; Optimized the XML skeleton.
''' </remarks>
Public Class Exporter
    Dim strRow As String
    Dim strAltRow As String
    Dim xmlDoc As XmlDocument
    Dim xmlRow As XmlElement
    Dim xmlCell As XmlElement
    Dim xmlData As XmlElement
    Dim xmlColumn As XmlElement()
    Dim xmlTitle As XmlElement()
    Dim xmlHeader As XmlElement
    Dim xmlTable As XmlElement
    Dim xmlWorkSheet As XmlElement
    Dim xmlWorkBook As XmlElement
    Dim strNS As String = "urn:schemas-microsoft-com:office:spreadsheet"
    Dim iNonDataRows As Integer = 0

    ''' <summary>
    ''' Load the pieces of XML spreadsheet (Excel) template, fill them with data and concatenate them to form a valid XML file.
    ''' </summary>
    ''' <param name="dtExp">Data to be exported</param>
    ''' <param name="objFmt">Instance of the ExcelFormater providing all format information</param>
    ''' <returns>The string of XML spreadsheet</returns>
    ''' <remarks>Excel spreadsheet can be saved as XML file. The principle of this method is to prepare the Excel with XML spreadsheet templates.</remarks>
    Public Function ExportToExcel(ByVal dtExp As DataTable, ByVal objFmt As Formater) As String
        ' Initiate the XML document
        InitDoc(objFmt)

        ' Get the column width
        If objFmt.ColumnWidth.Count > 0 Then
            xmlColumn = GetColumnWidth(objFmt, dtExp.Columns)
        End If

        ' Get the header template
        If objFmt.ShowHeader Then
            xmlHeader = GetHeader(dtExp.Columns)
            iNonDataRows = 1
        End If

        ' Get the report title
        If objFmt.SearchParameters.Count > 0 Then
            xmlTitle = GetTitle(objFmt.ReportName, objFmt, dtExp.Columns.Count - 1)
            iNonDataRows += xmlTitle.Length
        End If

        ' Fill data
        FillData(objFmt, dtExp, iNonDataRows)

        Return xmlDoc.InnerXml
    End Function

    ''' <summary>
    ''' Initiate the XML document by loading the XML skeleton with predefined format.
    ''' </summary>
    ''' <param name="objFmt">Instance of Formater</param>
    ''' <remarks></remarks>
    Private Sub InitDoc(ByVal objFmt As Formater)
        ' Get the XML skeleton template
        Dim objResMgr As New Resources.ResourceManager("ExcelExporter.Templates", Assembly.GetExecutingAssembly())
        Dim sbSkeleton As New StringBuilder
        sbSkeleton.Append(objResMgr.GetString("Skeleton"))

        ' Set format
        sbSkeleton.Replace("$FtNm$", objFmt.FontName)
        sbSkeleton.Replace("$FtFmly$", objFmt.FontFamily)
        sbSkeleton.Replace("$ClFtClr$", objFmt.CellFontColor)
        sbSkeleton.Replace("$ClBkClr$", objFmt.CellBackColor)
        sbSkeleton.Replace("$AltClBkClr$", objFmt.AlterCellBackColor)
        sbSkeleton.Replace("$DecFmt$", objFmt.DecimalFormat)
        sbSkeleton.Replace("$CurFmt$", objFmt.SpecificNumberFormat(SpecificNumber.Currency))
        sbSkeleton.Replace("$PerFmt$", objFmt.SpecificNumberFormat(SpecificNumber.Percentage))
        sbSkeleton.Replace("$FraFmt$", objFmt.SpecificNumberFormat(SpecificNumber.Fraction))
        sbSkeleton.Replace("$SciFmt$", objFmt.SpecificNumberFormat(SpecificNumber.Scientific))
        sbSkeleton.Replace("$DtFmt$", objFmt.DateTimeFormat)
        sbSkeleton.Replace("$HdFtClr$", objFmt.HeaderFontColor)
        sbSkeleton.Replace("$HdBkClr$", objFmt.HeaderBackColor)

        ' Initialize the XML document
        xmlDoc = New XmlDocument()
        xmlDoc.LoadXml(sbSkeleton.ToString)

        ' Initialize the element templates
        xmlRow = xmlDoc.CreateElement("Row", strNS)
        xmlCell = xmlDoc.CreateElement("Cell", strNS)
        xmlData = xmlDoc.CreateElement("Data", strNS)
        xmlWorkSheet = xmlDoc.CreateElement("Worksheet", strNS)
        xmlWorkBook = xmlDoc.GetElementsByTagName("Workbook")(0)   ' Only 1 workbook is included in a spreadsheet
        xmlTable = xmlDoc.CreateElement("Table", strNS)
        xmlTable.SetAttribute("FullColumns", strNS, "1")
        xmlTable.SetAttribute("FullRows", strNS, "1")
    End Sub

    ''' <summary>
    ''' Create the header template.
    ''' </summary>
    ''' <param name="dccExp">Collection of all the DataColumn in the DataTable to be exported</param>
    ''' <returns>XmlElement for the report header</returns>
    ''' <remarks></remarks>
    Private Function GetHeader(ByVal dccExp As DataColumnCollection) As XmlElement
        Dim xmlHdRow As XmlElement = xmlRow.Clone
        Dim xmlHdCell As XmlElement
        Dim xmlHdData As XmlElement
        For Each dcHeader As DataColumn In dccExp
            xmlHdCell = xmlCell.Clone
            xmlHdData = xmlData.Clone
            xmlHdCell.SetAttribute("StyleID", strNS, "sHeader")
            xmlHdData.SetAttribute("Type", strNS, "String")
            xmlHdData.InnerText = dcHeader.ColumnName
            xmlHdCell.AppendChild(xmlHdData)
            xmlHdRow.AppendChild(xmlHdCell)
        Next

        Return xmlHdRow
    End Function

    ''' <summary>
    ''' Get the report title template.
    ''' </summary>
    ''' <param name="strRptNm">Report name</param>
    ''' <param name="objFmt">Instance of the Formater</param>
    ''' <param name="iMerge">Number of columns to merge</param>
    ''' <returns>Array of XmlElement containing the rows for report name and search parameters</returns>
    ''' <remarks>The value of iMerge does not include the first cell of the merge range. For example, if the 3 columns merge into 1, then iMerge should be 2.</remarks>
    Private Function GetTitle(ByVal strRptNm As String, ByVal objFmt As Formater, ByVal iMerge As Integer) As XmlElement()
        Dim arlTitle As New ArrayList
        Dim xmlTlRow As XmlElement
        Dim xmlTlCell As XmlElement
        Dim xmlTlData As XmlElement

        ' Report name
        xmlTlRow = xmlRow.Clone
        xmlTlCell = xmlCell.Clone
        xmlTlData = xmlData.Clone

        xmlTlCell.SetAttribute("MergeAcross", strNS, iMerge)
        xmlTlCell.SetAttribute("StyleID", strNS, "sRptNm")
        xmlTlData.SetAttribute("Type", strNS, "String")

        xmlTlData.InnerText = strRptNm
        xmlTlCell.AppendChild(xmlTlData)
        xmlTlRow.AppendChild(xmlTlCell)

        arlTitle.Add(xmlTlRow)

        ' Parameters
        Dim sbPara As New StringBuilder
        For Each strKey As String In objFmt.SearchParameters.Keys
            xmlTlRow = xmlRow.Clone
            xmlTlCell = xmlCell.Clone
            xmlTlData = xmlData.Clone
            xmlTlCell.SetAttribute("MergeAcross", strNS, iMerge)
            xmlTlCell.SetAttribute("StyleID", strNS, "sPara")
            xmlTlData.SetAttribute("Type", strNS, "String")

            sbPara.Clear()
            sbPara.Append(strKey)
            sbPara.Append(" - ")
            sbPara.Append(objFmt.SearchParameters(strKey))

            xmlTlData.InnerText = sbPara.ToString
            xmlTlCell.AppendChild(xmlTlData)
            xmlTlRow.AppendChild(xmlTlCell)

            arlTitle.Add(xmlTlRow)
        Next

        Return arlTitle.ToArray(GetType(XmlElement))
    End Function

    ''' <summary>
    ''' Get the width of columns
    ''' </summary>
    ''' <param name="objFmt">Instance of Formater</param>
    ''' <param name="dcData">DataColumnCollection of the DataTable to be exported</param>
    ''' <returns>Array of Column XmlElement</returns>
    ''' <remarks></remarks>
    Private Function GetColumnWidth(ByVal objFmt As Formater, ByVal dcData As DataColumnCollection) As XmlElement()
        ' Sort the column element, the Index attribute must be in ascending sequence
        Dim iOrd As Integer
        Dim dtSort As New DataTable
        Dim drSort As DataRow
        dtSort.Columns.Add("Ordinal", GetType(System.Int32))
        dtSort.Columns.Add("Width")
        For Each strCol As String In objFmt.ColumnWidth.Keys
            iOrd = dcData(strCol).Ordinal
            If iOrd <> -1 Then
                drSort = dtSort.NewRow
                drSort("Ordinal") = iOrd + 1   ' The DataColumn ordinal is 0-based while ss:Index is 1-based.
                drSort("Width") = objFmt.ColumnWidth(strCol)
                dtSort.Rows.Add(drSort)
            End If
        Next
        Dim dvSort As New DataView(dtSort)
        dvSort.Sort = "Ordinal"
        dtSort = dvSort.ToTable

        ' Collect the column elements
        Dim xmlDtColumn As XmlElement
        Dim arlColWth As New ArrayList
        Dim xmlColumn As XmlElement = xmlDoc.CreateElement("Column", strNS)
        For Each drCol As DataRow In dtSort.Rows
            xmlDtColumn = xmlColumn.Clone
            xmlDtColumn.SetAttribute("Index", strNS, drCol("Ordinal"))
            xmlDtColumn.SetAttribute("Width", strNS, drCol("Width"))
            arlColWth.Add(xmlDtColumn)
        Next

        Return arlColWth.ToArray(GetType(XmlElement))
    End Function

    ''' <summary>
    ''' Fill the data.
    ''' </summary>
    ''' <param name="objFmt">Instance of Formater</param>
    ''' <param name="dtData">Data to be exported</param>
    ''' <param name="iNnDt">Number of non-data rows in each worksheet</param>
    ''' <remarks>Auto-pagination is equiped.</remarks>
    Private Sub FillData(ByVal objFmt As Formater, ByVal dtData As DataTable, ByVal iNnDt As Integer)
        ' Get row templates
        Dim xmlDtRow As XmlElement = xmlRow.Clone
        Dim xmlDtCell As XmlElement
        Dim xmlDtData As XmlElement
        Dim xmlAltDtRow As XmlElement = xmlRow.Clone
        Dim xmlAltDtCell As XmlElement
        Dim xmlAltDtData As XmlElement
        Dim xmlDtWorkSheet As XmlElement
        Dim xmlDtTable As XmlElement

        For Each dcData As DataColumn In dtData.Columns
            xmlDtCell = xmlCell.Clone
            xmlDtData = xmlData.Clone
            xmlAltDtCell = xmlCell.Clone
            xmlAltDtData = xmlData.Clone

            If objFmt.ColumnNumberFormat.ContainsKey(dcData.ColumnName) Then
                Select Case objFmt.ColumnNumberFormat(dcData.ColumnName)
                    Case SpecificNumber.Currency
                        xmlDtCell.SetAttribute("StyleID", strNS, "sCur")
                        xmlAltDtCell.SetAttribute("StyleID", strNS, "sCurAlt")
                        xmlDtData.SetAttribute("Type", strNS, "Number")
                        xmlAltDtData.SetAttribute("Type", strNS, "Number")
                    Case SpecificNumber.Percentage
                        xmlDtCell.SetAttribute("StyleID", strNS, "sPer")
                        xmlAltDtCell.SetAttribute("StyleID", strNS, "sPerAlt")
                        xmlDtData.SetAttribute("Type", strNS, "Number")
                        xmlAltDtData.SetAttribute("Type", strNS, "Number")
                    Case SpecificNumber.Fraction
                        xmlDtCell.SetAttribute("StyleID", strNS, "sFra")
                        xmlAltDtCell.SetAttribute("StyleID", strNS, "sFraAlt")
                        xmlDtData.SetAttribute("Type", strNS, "Number")
                        xmlAltDtData.SetAttribute("Type", strNS, "Number")
                    Case SpecificNumber.Scientific
                        xmlDtCell.SetAttribute("StyleID", strNS, "sSci")
                        xmlAltDtCell.SetAttribute("StyleID", strNS, "sSciAlt")
                        xmlDtData.SetAttribute("Type", strNS, "Number")
                        xmlAltDtData.SetAttribute("Type", strNS, "Number")
                End Select
            Else
                Select Case dcData.DataType
                    Case Type.GetType("System.Int16"), Type.GetType("System.Int32"), Type.GetType("System.Int64")
                        xmlDtCell.SetAttribute("StyleID", strNS, "sInt")
                        xmlAltDtCell.SetAttribute("StyleID", strNS, "sIntAlt")
                        xmlDtData.SetAttribute("Type", strNS, "Number")
                        xmlAltDtData.SetAttribute("Type", strNS, "Number")
                    Case Type.GetType("System.Decimal")
                        xmlDtCell.SetAttribute("StyleID", strNS, "sDec")
                        xmlAltDtCell.SetAttribute("StyleID", strNS, "sDecAlt")
                        xmlDtData.SetAttribute("Type", strNS, "Number")
                        xmlAltDtData.SetAttribute("Type", strNS, "Number")
                    Case Type.GetType("System.DateTime")
                        xmlDtCell.SetAttribute("StyleID", strNS, "sDT")
                        xmlAltDtCell.SetAttribute("StyleID", strNS, "sDTAlt")
                        xmlDtData.SetAttribute("Type", strNS, "DateTime")
                        xmlAltDtData.SetAttribute("Type", strNS, "DateTime")
                    Case Else
                        If objFmt.CenterAlignColumns.Contains(dcData.ColumnName) Then
                            xmlDtCell.SetAttribute("StyleID", strNS, "sStrCtr")
                            xmlAltDtCell.SetAttribute("StyleID", strNS, "sStrCtrAlt")
                            xmlDtData.SetAttribute("Type", strNS, "String")
                            xmlAltDtData.SetAttribute("Type", strNS, "String")
                        Else
                            xmlDtCell.SetAttribute("StyleID", strNS, "sStr")
                            xmlAltDtCell.SetAttribute("StyleID", strNS, "sStrAlt")
                            xmlDtData.SetAttribute("Type", strNS, "String")
                            xmlAltDtData.SetAttribute("Type", strNS, "String")
                        End If
                End Select
            End If
            xmlDtCell.AppendChild(xmlDtData)
            xmlDtRow.AppendChild(xmlDtCell)
            xmlAltDtCell.AppendChild(xmlAltDtData)
            xmlAltDtRow.AppendChild(xmlAltDtCell)
        Next

        ' Calculate the number of worksheets
        Dim iDataRows As Integer = objFmt.MaxRowPerSheet - iNnDt
        Dim iPage As Integer = dtData.Rows.Count \ iDataRows   ' Cut the decimal
        If dtData.Rows.Count Mod iDataRows = 0 Then   ' The page count is 0-based.
            iPage -= 1
        End If

        ' Fill each worksheet with data
        For iCount As Integer = 0 To iPage
            Dim sbSheetName As New StringBuilder
            Dim iStrIdx As Integer = iCount * iDataRows
            Dim iEndIdx As Integer
            If iCount = iPage Then
                iEndIdx = dtData.Rows.Count - 1
            Else
                iEndIdx = iStrIdx + iDataRows - 1
            End If

            sbSheetName.Append(objFmt.ReportName)
            If iPage > 0 Then
                sbSheetName.Append(" - ")
                sbSheetName.Append(iCount + 1)
            End If

            xmlDtWorkSheet = xmlWorkSheet.Clone
            xmlDtTable = xmlTable.Clone
            xmlDtWorkSheet.SetAttribute("Name", strNS, sbSheetName.ToString)
            xmlDtTable.SetAttribute("ExpandedColumnCount", strNS, dtData.Columns.Count)
            xmlDtTable.SetAttribute("ExpandedRowCount", strNS, iEndIdx - iStrIdx + 1 + iNonDataRows)

            ' Set column width
            If Not xmlColumn Is Nothing Then
                For iCol As Integer = 0 To xmlColumn.Length - 1
                    xmlDtTable.AppendChild(xmlColumn(iCol).Clone)
                Next
            End If

            ' Insert title section
            If Not xmlTitle Is Nothing Then
                For iTitle As Integer = 0 To xmlTitle.Length - 1
                    xmlDtTable.AppendChild(xmlTitle(iTitle).Clone)
                Next
            End If

            ' Insert header row
            If Not xmlHeader Is Nothing Then
                xmlDtTable.AppendChild(xmlHeader.Clone)
            End If

            Dim bAlter As Boolean = False
            Dim xmlRowData As XmlElement
            Dim drData As DataRow
            Dim strData As String
            For iRow As Integer = iStrIdx To iEndIdx
                drData = dtData.Rows(iRow)
                If bAlter Then
                    xmlRowData = xmlAltDtRow.Clone
                Else
                    xmlRowData = xmlDtRow.Clone
                End If

                For iCol As Integer = 0 To dtData.Columns.Count - 1
                    strData = drData(iCol).ToString

                    If strData <> String.Empty Then
                        If dtData.Columns(iCol).DataType = Type.GetType("System.DateTime") Then
                            strData = Format(CDate(strData), "yyyy-MM-ddThh:mm:ss.sss")
                        End If
                        xmlRowData.ChildNodes(iCol).ChildNodes(0).InnerText = strData
                    Else
                        xmlRowData.ChildNodes(iCol).RemoveChild(xmlRowData.ChildNodes(iCol).ChildNodes(0))
                    End If
                Next

                xmlDtTable.AppendChild(xmlRowData)
                bAlter = Not bAlter
            Next

            xmlDtWorkSheet.AppendChild(xmlDtTable)
            xmlWorkBook.AppendChild(xmlDtWorkSheet)
        Next
    End Sub
End Class