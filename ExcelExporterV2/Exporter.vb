Imports System.Text
Imports System.Reflection

''' <summary>
''' This class takes in data in the type of DataTable and output a string in XML spreadsheet format.
''' </summary>
''' <remarks>
''' Created by Alex Liang on 4-Apr-2011
''' V2.0 upgrade by Alex Liang on 4-Nov-2011
''' </remarks>
Public Class Exporter
    Dim strRow As String
    Dim strAltRow As String

    ''' <summary>
    ''' Load the pieces of XML spreadsheet (Excel) template, fill them with data and concatenate them to form a valid XML file.
    ''' </summary>
    ''' <param name="dtExp">Data to be exported</param>
    ''' <param name="objFormater">Instance of the ExcelFormater providing all format information</param>
    ''' <returns>The string of XML spreadsheet</returns>
    ''' <remarks>Excel spreadsheet can be saved as XML file. The principle of this method is to prepare the Excel with XML spreadsheet templates.</remarks>
    Public Function ExportToExcel(ByVal dtExp As DataTable, ByVal objFormater As Formater) As String
        Dim objResMgr As New Resources.ResourceManager("ExcelExporter.Templates", Assembly.GetExecutingAssembly())
        Dim strHead As String = objResMgr.GetString("Head")
        Dim strSheetHead As String = objResMgr.GetString("SheetHead")
        Dim strCellHead As String = objResMgr.GetString("CellHead")
        Dim strCellTail As String = objResMgr.GetString("CellTail")
        Dim strSheetTail As String = objResMgr.GetString("SheetTail")
        Dim strTail As String = objResMgr.GetString("Tail")
        Dim sbXML As New StringBuilder
        sbXML.Append(strHead)

        ' Implement the format settings
        sbXML.Replace("$FtNm$", objFormater.FontName)
        sbXML.Replace("$FtFmly$", objFormater.FontFamily)
        sbXML.Replace("$ClFtClr$", objFormater.CellFontColor)
        sbXML.Replace("$ClBkClr$", objFormater.CellBackColor)
        sbXML.Replace("$AltClBkClr$", objFormater.AlterCellBackColor)
        sbXML.Replace("$DecFmt$", objFormater.DecimalFormat)
        sbXML.Replace("$DtFmt$", objFormater.DateTimeFormat)
        sbXML.Replace("$HdFtClr$", objFormater.HeaderFontColor)
        sbXML.Replace("$HdBkClr$", objFormater.HeaderBackColor)

        ' Initialize header and prepare row template
        Dim sbHeader As New StringBuilder
        Dim sbRow As New StringBuilder
        Dim sbAltRow As New StringBuilder
        Dim sbHolder As New StringBuilder
        Dim iIndex As Integer = 0
        Dim strHeaderHead As String = strCellHead.Replace("$Style$", "sHeader").Replace("$Type$", "String")
        sbHeader.Append("<Row>")
        sbHeader.Append(Constants.vbNewLine)
        sbRow.Append("<Row>")
        sbRow.Append(Constants.vbNewLine)
        sbAltRow.Append("<Row>")
        sbAltRow.Append(Constants.vbNewLine)
        For Each dcExp As DataColumn In dtExp.Columns
            sbHeader.Append(strHeaderHead)
            sbHeader.Append(dcExp.ColumnName)
            sbHeader.Append(strCellTail)
            sbHeader.Append(Constants.vbNewLine)

            sbHolder.Clear()
            sbHolder.Append("$")
            sbHolder.Append(iIndex)
            sbHolder.Append("$")
            Select Case dcExp.DataType
                Case Type.GetType("System.Int16"), Type.GetType("System.Int32"), Type.GetType("System.Int64")
                    sbRow.Append(strCellHead.Replace("$Style$", "sInt").Replace("$Type$", "Number"))
                    sbAltRow.Append(strCellHead.Replace("$Style$", "sIntAlt").Replace("$Type$", "Number"))
                Case Type.GetType("System.Decimal")
                    sbRow.Append(strCellHead.Replace("$Style$", "sDec").Replace("$Type$", "Number"))
                    sbAltRow.Append(strCellHead.Replace("$Style$", "sDecAlt").Replace("$Type$", "Number"))
                Case Type.GetType("System.DateTime")
                    sbRow.Append(strCellHead.Replace("$Style$", "sDT").Replace("$Type$", "DateTime"))
                    sbAltRow.Append(strCellHead.Replace("$Style$", "sDTAlt").Replace("$Type$", "DateTime"))
                Case Else
                    If objFormater.CenterAlignColumns.Contains(dcExp.ColumnName) Then
                        sbRow.Append(strCellHead.Replace("$Style$", "sStrCtr").Replace("$Type$", "String"))
                        sbAltRow.Append(strCellHead.Replace("$Style$", "sStrCtrAlt").Replace("$Type$", "String"))
                    Else
                        sbRow.Append(strCellHead.Replace("$Style$", "sStr").Replace("$Type$", "String"))
                        sbAltRow.Append(strCellHead.Replace("$Style$", "sStrAlt").Replace("$Type$", "String"))
                    End If
            End Select
            sbRow.Append(sbHolder.ToString)
            sbRow.Append(strCellTail)
            sbRow.Append(Constants.vbNewLine)
            sbAltRow.Append(sbHolder.ToString)
            sbAltRow.Append(strCellTail)
            sbAltRow.Append(Constants.vbNewLine)
            iIndex += 1
        Next
        sbHeader.Append("</Row>")
        sbHeader.Append(Constants.vbNewLine)
        sbRow.Append("</Row>")
        sbRow.Append(Constants.vbNewLine)
        sbAltRow.Append("</Row>")
        sbAltRow.Append(Constants.vbNewLine)
        strRow = sbRow.ToString
        strAltRow = sbAltRow.ToString

        Dim iDataRows As Integer = 65536   ' Excel worksheet can hold 65536 rows in maximum.
        Dim strRptName As String = String.Empty
        Dim iNonDataRows As Integer = 0

        If objFormater.ShowHeader Then
            iNonDataRows = 1   ' The default number of nondata row is one, which is for the report header.
        End If

        If Not objFormater.SearchParameters Is Nothing Then
            iNonDataRows += (objFormater.SearchParameters.Count + 1)   ' If it's going to show the title, count the rows occupied by the report name and search parameters on nondata rows.
            Dim sbClHd As New StringBuilder
            sbClHd.Append(objResMgr.GetString("TitleCellHead"))
            Dim iMerge As Integer = dtExp.Columns.Count - 1
            Dim sbRptName As New StringBuilder
            sbRptName.Append(Constants.vbNewLine)
            sbRptName.Append("<Row>")
            sbRptName.Append(Constants.vbNewLine)
            sbRptName.Append(sbClHd.Replace("$Cols$", iMerge).Replace("$Style$", "sRptNm").ToString)
            sbRptName.Append(objFormater.ReportName)
            sbRptName.Append(strCellTail)
            sbRptName.Append(Constants.vbNewLine)
            sbRptName.Append("</Row>")
            sbRptName.Append(Constants.vbNewLine)

            Dim strPara As String = sbClHd.Replace("sRptNm", "sPara").ToString
            For Each objPara As DictionaryEntry In objFormater.SearchParameters
                sbRptName.Append("<Row>")
                sbRptName.Append(Constants.vbNewLine)
                sbRptName.Append(strPara)
                sbRptName.Append(objPara.Key.ToString)
                sbRptName.Append(" - ")
                sbRptName.Append(objPara.Value.ToString)
                sbRptName.Append(strCellTail)
                sbRptName.Append(Constants.vbNewLine)
                sbRptName.Append("</Row>")
                sbRptName.Append(Constants.vbNewLine)
            Next

            strRptName = sbRptName.ToString
        End If

        iDataRows -= iNonDataRows   ' Exclude all the nondata rows from the default value of data rows.

        Dim iPage As Integer = dtExp.Rows.Count \ iDataRows
        If dtExp.Rows.Count Mod iDataRows = 0 Then
            iPage = iPage - 1
        End If
        For iCount As Integer = 0 To iPage
            Dim sbWorkSheet As New StringBuilder
            Dim sbSheetName As New StringBuilder
            Dim iStrIdx As Integer = iCount * iDataRows
            Dim iEndIdx As Integer
            If iCount = iPage Then
                iEndIdx = dtExp.Rows.Count - 1
            Else
                iEndIdx = iStrIdx + iDataRows - 1
            End If

            sbSheetName.Append(objFormater.ReportName)
            sbSheetName.Append(" - ")
            sbSheetName.Append(iCount + 1)
            sbWorkSheet.Append(strSheetHead)
            sbWorkSheet.Replace("$SheetName$", sbSheetName.ToString)
            sbWorkSheet.Replace("$ColCount$", dtExp.Columns.Count)
            sbWorkSheet.Replace("$RowCount$", iEndIdx - iStrIdx + 1 + iNonDataRows)
            sbXML.Append(sbWorkSheet.ToString)
            If strRptName <> String.Empty Then
                sbXML.Append(strRptName)
            End If
            If objFormater.ShowHeader Then
                sbXML.Append(sbHeader.ToString)
            End If
            AttachRows(dtExp, iStrIdx, iEndIdx, sbXML)
            sbXML.Append(strSheetTail)
        Next

        sbXML.Replace("<Data ss:Type=""DateTime""></Data>", "")   ' Remove the empty dates
        sbXML.Append(Constants.vbNewLine)
        sbXML.Append(strTail)
        Return sbXML.ToString
    End Function

    ''' <summary>
    ''' Fill data to each row with the XML row template.
    ''' </summary>
    ''' <param name="dtRow">DataTable containing the data of all the rows to be exported</param>
    ''' <param name="iStrIdx">DataTable index of the first row in the current page</param>
    ''' <param name="iEndIdx">DataTable index of the last row in the current page</param>
    ''' <param name="sbXML">StringBuilder concantenating each XML piece templates</param>
    ''' <remarks>It fills data for one single page of the spreadsheet in each invoke.</remarks>
    Private Sub AttachRows(ByVal dtRow As DataTable, ByVal iStrIdx As Integer, ByVal iEndIdx As Integer, ByVal sbXML As StringBuilder)
        Dim bAlter As Boolean = False
        Dim sbHolder As New StringBuilder
        For iCount As Integer = iStrIdx To iEndIdx
            Dim drRow As DataRow = dtRow.Rows(iCount)
            Dim sbRow As New StringBuilder
            If bAlter Then
                sbRow.Append(strAltRow)
            Else
                sbRow.Append(strRow)
            End If

            For iCol As Integer = 0 To dtRow.Columns.Count - 1
                sbHolder.Clear()
                sbHolder.Append("$")
                sbHolder.Append(iCol)
                sbHolder.Append("$")

                Dim strRow As String = drRow(iCol).ToString

                If dtRow.Columns(iCol).DataType = Type.GetType("System.DateTime") AndAlso strRow <> String.Empty Then
                    sbRow.Replace(sbHolder.ToString, Format(CDate(strRow), "yyyy-MM-ddThh:mm:ss.sss"))
                Else
                    sbRow.Replace(sbHolder.ToString, strRow)
                End If
            Next

            sbXML.Append(sbRow.ToString)
            bAlter = Not bAlter
        Next
    End Sub
End Class