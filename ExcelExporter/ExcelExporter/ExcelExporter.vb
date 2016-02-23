Imports System.Text
Imports System.Reflection

''' <summary>
''' This class takes in data in the type of DataTable and output a string in XML spreadsheet format.
''' </summary>
''' <remarks>
''' Created by Alex Liang on 4-Apr-2011
''' </remarks>
Public Class ExcelExporter
    Dim strRow As String
    Dim strAltRow As String

    ''' <summary>
    ''' Load the pieces of XML spreadsheet (Excel) template, fill them with data and concatenate them to form a valid XML file.
    ''' </summary>
    ''' <param name="dtExp">DataTable to be exported</param>
    ''' <returns>The string of XML spreadsheet</returns>
    ''' <remarks>Excel spreadsheet can be saved as XML file. The principle of this method is to prepare the Excel with XML spreadsheet templates.</remarks>
    Public Function ExportToExcel(ByVal dtExp As DataTable, ByVal strReportName As String, ByVal arrCenAlgStr As Integer()) As String
        Dim objResMgr As New Resources.ResourceManager("ExcelExporter.Templates", Assembly.GetExecutingAssembly())
        Dim strHead As String = objResMgr.GetString("Head")
        Dim strSheetHead As String = objResMgr.GetString("SheetHead")
        Dim strCellHead As String = objResMgr.GetString("CellHead")
        Dim strCellTail As String = objResMgr.GetString("CellTail")
        Dim strSheetTail As String = objResMgr.GetString("SheetTail")
        Dim strTail As String = objResMgr.GetString("Tail")
        Dim sbXML As New StringBuilder
        sbXML.Append(strHead)

        ' Initialize header and prepare row template
        Dim sbHeader As New StringBuilder
        Dim sbRow As New StringBuilder
        Dim sbAltRow As New StringBuilder
        Dim sbHolder As New StringBuilder
        Dim iIndex As Integer = 0
        Dim strHeaderHead As String = strCellHead.Replace("$Style$", "s32").Replace("$Type$", "String")
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
                    sbRow.Append(strCellHead.Replace("$Style$", "s24").Replace("$Type$", "Number"))
                    sbAltRow.Append(strCellHead.Replace("$Style$", "s29").Replace("$Type$", "Number"))
                Case Type.GetType("System.Decimal")
                    sbRow.Append(strCellHead.Replace("$Style$", "s25").Replace("$Type$", "Number"))
                    sbAltRow.Append(strCellHead.Replace("$Style$", "s30").Replace("$Type$", "Number"))
                Case Type.GetType("System.DateTime")
                    sbRow.Append(strCellHead.Replace("$Style$", "s26").Replace("$Type$", "DateTime"))
                    sbAltRow.Append(strCellHead.Replace("$Style$", "s31").Replace("$Type$", "DateTime"))
                Case Else
                    If arrCenAlgStr.Contains(iIndex) Then
                        sbRow.Append(strCellHead.Replace("$Style$", "s22").Replace("$Type$", "String"))
                        sbAltRow.Append(strCellHead.Replace("$Style$", "s27").Replace("$Type$", "String"))
                    Else
                        sbRow.Append(strCellHead.Replace("$Style$", "s23").Replace("$Type$", "String"))
                        sbAltRow.Append(strCellHead.Replace("$Style$", "s28").Replace("$Type$", "String"))
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

        Dim iPage As Integer = dtExp.Rows.Count \ 65535
        If dtExp.Rows.Count Mod 65535 = 0 Then
            iPage = iPage - 1
        End If
        For iCount As Integer = 0 To iPage
            Dim sbWorkSheet As New StringBuilder
            Dim sbSheetName As New StringBuilder
            Dim iStrIdx As Integer = iCount * 65535
            Dim iEndIdx As Integer
            If iCount = iPage Then
                iEndIdx = dtExp.Rows.Count - 1
            Else
                iEndIdx = iStrIdx + 65534
            End If

            sbSheetName.Append(strReportName)
            sbSheetName.Append(" - ")
            sbSheetName.Append(iCount + 1)
            sbWorkSheet.Append(strSheetHead)
            sbWorkSheet.Replace("$SheetName$", sbSheetName.ToString)
            sbWorkSheet.Replace("$ColCount$", dtExp.Columns.Count)
            sbWorkSheet.Replace("$RowCount$", iEndIdx - iStrIdx + 2)
            sbXML.Append(sbWorkSheet.ToString)
            sbXML.Append(sbHeader.ToString)
            AttachRows(dtExp, iStrIdx, iEndIdx, sbXML)
            sbXML.Append(strSheetTail)
        Next

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
                If dtRow.Columns(iCol).DataType = Type.GetType("System.DataTime") Then
                    sbRow.Replace(sbHolder.ToString, Format(CDate(drRow(iCol).ToString), "yyyy-MM-dd"))
                Else
                    sbRow.Replace(sbHolder.ToString, drRow(iCol).ToString)
                End If
            Next

            sbXML.Append(sbRow.ToString)
            bAlter = Not bAlter
        Next
    End Sub
End Class
