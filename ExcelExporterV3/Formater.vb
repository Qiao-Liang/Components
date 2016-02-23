''' <summary>
''' This class helps setting the format of the Excel to be exported.
''' </summary>
''' <remarks>
''' Created by Alex Liang on 4-Nov-2011
''' V3.0 upgrade by Alex Liang on 16-Nov-2011
''' V3.1 upgrade by Alex Liang on 14-Dec-2011 -- Replaced the Hashtable to Dictionary as the output sequence of Hashtable keys is inconsistent with the input
''' V3.2 upgrade by Alex Liang on 11-Jan-2012 -- Introduced more number formats; Optimized the XML skeleton.
''' </remarks>
Public Class Formater
    Private iMaxRow As Integer
    Private strFtNm As String
    Private strFtFmly As String
    Private strRptNm As String
    Private strClFtClr As String
    Private strClBkClr As String
    Private strAltClBkClr As String
    Private strHdFtClr As String
    Private strHdBkClr As String
    Private strDecFmt As String
    Private strDtFmt As String
    Private bShwHdr As String
    Private strCtrAlgClmn() As String
    Private dicSrhPrs As Dictionary(Of String, String)
    Private dicColWth As Dictionary(Of String, Integer)
    Private dicNmFmt As Dictionary(Of SpecificNumber, String)
    Private dicColFmt As Dictionary(Of String, SpecificNumber)

    ''' <summary>
    ''' Enumeration of the specific number formats.
    ''' </summary>
    ''' <remarks>Introduced in the V3.1 upgrade.</remarks>
    Public Enum SpecificNumber
        Currency = 0
        Percentage = 1
        Fraction = 2
        Scientific = 3
    End Enum

    ''' <summary>
    ''' Defines the format of each specific type of number.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Introduced in the V3.1 upgrade.</remarks>
    Public Property SpecificNumberFormat As Dictionary(Of SpecificNumber, String)
        Get
            Return dicNmFmt
        End Get
        Set(ByVal value As Dictionary(Of SpecificNumber, String))
            dicNmFmt = value
        End Set
    End Property

    ''' <summary>
    ''' Specifies the format of columns in number type.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Introduced in the V3.1 upgrade. This is not mandatory. The columns, which is in number type but not specified here, will be rendered into basic number format.</remarks>
    Public Property ColumnNumberFormat As Dictionary(Of String, SpecificNumber)
        Get
            Return dicColFmt
        End Get
        Set(ByVal value As Dictionary(Of String, SpecificNumber))
            dicColFmt = value
        End Set
    End Property

    ''' <summary>
    ''' Maximum row count in a worksheet.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The default value is 65536, which is the maximum of Excel 2003.</remarks>
    Public Property MaxRowPerSheet As Integer
        Get
            Return iMaxRow
        End Get
        Set(ByVal value As Integer)
            iMaxRow = value
        End Set
    End Property

    ''' <summary>
    ''' Font name.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, Calibri</remarks>
    Public Property FontName As String
        Get
            Return strFtNm
        End Get
        Set(ByVal value As String)
            strFtNm = value
        End Set
    End Property

    ''' <summary>
    ''' Font family.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, Swiss</remarks>
    Public Property FontFamily As String
        Get
            Return strFtFmly
        End Get
        Set(ByVal value As String)
            strFtFmly = value
        End Set
    End Property

    ''' <summary>
    ''' The name of the report.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Can be taken as the sheet name and the name of the report.</remarks>
    Public Property ReportName As String
        Get
            Return strRptNm
        End Get
        Set(ByVal value As String)
            strRptNm = value
        End Set
    End Property

    ''' <summary>
    ''' Hex color of cell font.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, #000000</remarks>
    Public Property CellFontColor As String
        Get
            Return strClFtClr
        End Get
        Set(ByVal value As String)
            strClFtClr = value
        End Set
    End Property

    ''' <summary>
    ''' Hex color of cell background.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, #FFFFFF</remarks>
    Public Property CellBackColor As String
        Get
            Return strClBkClr
        End Get
        Set(ByVal value As String)
            strClBkClr = value
        End Set
    End Property

    ''' <summary>
    ''' Hex color of alternative cell background.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, #C0C0C0</remarks>
    Public Property AlterCellBackColor As String
        Get
            Return strAltClBkClr
        End Get
        Set(ByVal value As String)
            strAltClBkClr = value
        End Set
    End Property

    ''' <summary>
    ''' Hex color of header font.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, #FFFFFF</remarks>
    Public Property HeaderFontColor As String
        Get
            Return strHdFtClr
        End Get
        Set(ByVal value As String)
            strHdFtClr = value
        End Set
    End Property

    ''' <summary>
    ''' Hex color of header background. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, #3366FF</remarks>
    Public Property HeaderBackColor As String
        Get
            Return strHdBkClr
        End Get
        Set(ByVal value As String)
            strHdBkClr = value
        End Set
    End Property

    ''' <summary>
    ''' Format of Decimal.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, 0.00000</remarks>
    Public Property DecimalFormat As String
        Get
            Return strDecFmt
        End Get
        Set(ByVal value As String)
            strDecFmt = value
        End Set
    End Property

    ''' <summary>
    ''' Format of DateTime columns.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>For example, dd/MM/yyyy</remarks>
    Public Property DateTimeFormat As String
        Get
            Return strDtFmt
        End Get
        Set(ByVal value As String)
            strDtFmt = value
        End Set
    End Property

    ''' <summary>
    ''' Decides whether show the report header or not.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The default value is True.</remarks>
    Public Property ShowHeader As Boolean
        Get
            Return bShwHdr
        End Get
        Set(ByVal value As Boolean)
            bShwHdr = value
        End Set
    End Property

    ''' <summary>
    ''' The array of names of the center-aligned columns
    ''' </summary>
    ''' <value>Array of String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CenterAlignColumns As String()
        Get
            Return strCtrAlgClmn
        End Get
        Set(ByVal value As String())
            strCtrAlgClmn = value
        End Set
    End Property

    ''' <summary>
    ''' Key/Value pairs of the search parameters
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SearchParameters As Dictionary(Of String, String)
        Get
            Return dicSrhPrs
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            dicSrhPrs = value
        End Set
    End Property

    ''' <summary>
    ''' Key/Value pairs of the column width
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ColumnWidth As Dictionary(Of String, Integer)
        Get
            Return dicColWth
        End Get
        Set(ByVal value As Dictionary(Of String, Integer))
            dicColWth = value
        End Set
    End Property

    ''' <summary>
    ''' Constructor.
    ''' </summary>
    ''' <remarks>Set the default values.</remarks>
    Public Sub New()
        iMaxRow = 65536
        strFtNm = "Calibri"
        strFtFmly = "Swiss"
        strClFtClr = "#000000"
        strClBkClr = "#FFFFFF"
        strAltClBkClr = "#C0C0C0"
        strHdFtClr = "#FFFFFF"
        strHdBkClr = "#3366FF"
        strDecFmt = "0.00000"
        strDtFmt = "Medium Date"
        bShwHdr = True
        strCtrAlgClmn = {""}
        dicSrhPrs = New Dictionary(Of String, String)
        dicColWth = New Dictionary(Of String, Integer)
        dicNmFmt = New Dictionary(Of SpecificNumber, String)
        dicNmFmt.Add(SpecificNumber.Currency, "&quot;$&quot;#,##0.00")
        dicNmFmt.Add(SpecificNumber.Percentage, "Percent")
        dicNmFmt.Add(SpecificNumber.Fraction, "#\ ?/?")
        dicNmFmt.Add(SpecificNumber.Scientific, "Scientific")
        dicColFmt = New Dictionary(Of String, SpecificNumber)
    End Sub
End Class