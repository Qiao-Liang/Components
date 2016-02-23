Imports System.Text
Imports System.Xml
Imports ExcelExporter

Public Class Customizer

    Private strPath As String
    Private arlFlds As ArrayList
    Private dsSrc As DataSet
    Private arrFmt As Array
    Dim strNS As String = "urn:schemas-microsoft-com:office:spreadsheet"

    ''' <summary>
    ''' Path to the template in OOXML format.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TemplatePath As String
        Get
            Return strPath
        End Get
        Set(ByVal value As String)
            strPath = value
        End Set
    End Property

    ''' <summary>
    ''' Field values to be filled to the template.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FieldValues As ArrayList
        Get
            Return arlFlds
        End Get
        Set(ByVal value As ArrayList)
            arlFlds = value
        End Set
    End Property

    ''' <summary>
    ''' Data source of the tables.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TableDataSources As DataSet
        Get
            Return dsSrc
        End Get
        Set(ByVal value As DataSet)
            dsSrc = value
        End Set
    End Property

    ''' <summary>
    ''' Array of Formater instances respectively associated with the DataTables in the DataSet as the data source.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TableFormaters As Array
        Get
            Return arrFmt
        End Get
        Set(ByVal value As Array)
            arrFmt = value
        End Set
    End Property

    ''' <summary>
    ''' Get the OOXML output.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExportToExcel() As String
        Dim sbLocator As New StringBuilder

        Dim xmlDoc As New XmlDocument
        xmlDoc.Load(Me.strPath)

        Dim xmlNSMgr As New XmlNamespaceManager(xmlDoc.NameTable)
        xmlNSMgr.AddNamespace("ss", strNS)
        Dim xmlTable As XmlElement = xmlDoc.SelectSingleNode("/ss:Workbook/ss:Worksheet/ss:Table", xmlNSMgr)
        Dim xmlFldLoc As XmlElement

        For iCount As Integer = 0 To Me.arlFlds.Count - 1
            sbLocator.Clear()
            sbLocator.Append("/ss:Workbook/ss:Worksheet/ss:Table/ss:Row/ss:Cell/ss:Data[text()='!Field_")
            sbLocator.Append(iCount)
            sbLocator.Append("']")

            xmlFldLoc = xmlDoc.SelectSingleNode(sbLocator.ToString, xmlNSMgr)
            xmlFldLoc.InnerText = arlFlds(iCount).ToString

            Select Case arlFlds(iCount).GetType.ToString
                Case "System.String"
                    xmlFldLoc.SetAttribute("Type", strNS, "String")
                Case "System.Int16", "System.Int32", "System.Int64"
                    xmlFldLoc.SetAttribute("Type", strNS, "Number")
                Case "System.Decimal"
                    xmlFldLoc.SetAttribute("Type", strNS, "Number")
                Case "System.DateTime"
                    xmlFldLoc.InnerText = Format(CDate(arlFlds(iCount).ToString), "yyyy-MM-ddThh:mm:ss.sss")   ' Regulate the DateTime format.
                    xmlFldLoc.SetAttribute("Type", strNS, "DateTime")
            End Select
        Next

        Dim xmlTblLoc As XmlElement
        Dim iRowIdx As Integer = 0
        Dim iCellIdx As Integer = 1
        Dim iCurBdrIdx As Integer   ' This captures the index of the right most cell of current table.
        Dim iMaxCellIdx As Integer = xmlTable.Attributes("ExpandedColumnCount", strNS).Value   ' This captures the index of the right most cell, which would be set to the worksheet attribute "ExpandedColumnCount".
        Dim iMaxRowIdx As Integer = 1

        If Me.dsSrc.Tables.Count > 0 Then
            For iCount As Integer = 0 To Me.dsSrc.Tables.Count - 1
                sbLocator.Clear()
                sbLocator.Append("/ss:Workbook/ss:Worksheet/ss:Table/ss:Row/ss:Cell/ss:Data[text()='!Table_")
                sbLocator.Append(iCount)
                sbLocator.Append("']")

                xmlTblLoc = xmlDoc.SelectSingleNode(sbLocator.ToString, xmlNSMgr)
                If Not xmlTblLoc.ParentNode.Attributes("Index", strNS) Is Nothing Then
                    iCellIdx = xmlTblLoc.ParentNode.Attributes("Index", strNS).Value
                End If

                If Not xmlTblLoc.ParentNode.ParentNode.Attributes("Index", strNS) Is Nothing Then
                    iRowIdx = xmlTblLoc.ParentNode.ParentNode.Attributes("Index", strNS).Value
                End If

                iCurBdrIdx = iCellIdx + dsSrc.Tables(iCount).Columns.Count - 1

                If iCurBdrIdx > iMaxCellIdx Then
                    iMaxCellIdx = iCurBdrIdx
                End If

                Dim xmlExpDoc As New XmlDocument
                Dim objExp As New Exporter
                xmlExpDoc.LoadXml(objExp.ExportToExcel(dsSrc.Tables(iCount), Me.arrFmt(iCount)))

                ' Add the style for the tables. Just add the style definition for one time since it's the same for all the tables.
                If iCount = 0 Then
                    Dim xmlStyles As XmlElement = xmlDoc.SelectSingleNode("/ss:Workbook/ss:Styles", xmlNSMgr)
                    Dim xmlExpStyles As XmlNodeList = xmlExpDoc.SelectNodes("/ss:Workbook/ss:Styles/ss:Style", xmlNSMgr)

                    For Each xmlStyle As XmlElement In xmlExpStyles
                        xmlStyles.AppendChild(xmlDoc.ImportNode(xmlStyle, True))
                    Next
                Else
                    For iTblCount As Integer = 0 To iCount - 1
                        iRowIdx += dsSrc.Tables(iTblCount).Rows.Count

                        If Not CType(arrFmt(iTblCount), Formater).ShowHeader Then   ' The table header takes over the row of the table locator, in this case, if no header appears, the number of increased row should be 1 less.
                            iRowIdx -= 1
                        End If
                    Next

                    If iCount = dsSrc.Tables.Count - 1 Then   ' Check the last table
                        iMaxRowIdx = iRowIdx + dsSrc.Tables(iCount).Rows.Count

                        If Not CType(arrFmt(iCount), Formater).ShowHeader Then   ' The table header takes over the row of the table locator, in this case, if no header appears, the number of increased row should be 1 less.
                            iMaxRowIdx -= 1
                        End If
                    End If
                End If

                Dim xmlExpTable As XmlElement = xmlExpDoc.SelectSingleNode("/ss:Workbook/ss:Worksheet/ss:Table", xmlNSMgr)
                If iRowIdx > 1 Then
                    CType(xmlExpTable.FirstChild, XmlElement).SetAttribute("Index", strNS, iRowIdx)
                End If

                If iCellIdx > 1 Then
                    For Each xmlRow As XmlElement In xmlExpTable.ChildNodes
                        CType(xmlRow.FirstChild, XmlElement).SetAttribute("Index", strNS, iCellIdx)
                    Next
                End If

                ' Add the table to the final XMLDocument.
                For Each xmlRows As XmlElement In xmlExpTable.ChildNodes
                    xmlTable.InsertBefore(xmlDoc.ImportNode(xmlRows, True), xmlTblLoc.ParentNode.ParentNode)
                Next

                xmlTable.RemoveChild(xmlTblLoc.ParentNode.ParentNode)   ' Remove the locator.
            Next

            xmlTable.SetAttribute("ExpandedColumnCount", strNS, iMaxCellIdx)
            xmlTable.SetAttribute("ExpandedRowCount", strNS, iMaxRowIdx)
        End If

        Return xmlDoc.InnerXml
    End Function
End Class
