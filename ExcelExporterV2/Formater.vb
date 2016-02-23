''' <summary>
''' This class helps setting the format of the Excel to be exported.
''' </summary>
''' <remarks>Created by Alex Liang on 4-Nov-2011</remarks>
Public Class Formater
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
    Private objSrhPrsHash As Hashtable

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
    Public Property SearchParameters As Hashtable
        Get
            Return objSrhPrsHash
        End Get
        Set(ByVal value As Hashtable)
            objSrhPrsHash = value
        End Set
    End Property

    ''' <summary>
    ''' Constructor.
    ''' </summary>
    ''' <remarks>Set the default values.</remarks>
    Public Sub New()
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
    End Sub
End Class