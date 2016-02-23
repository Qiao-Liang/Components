Imports System.Xml
''' <summary>
''' This class helps setting the format of the Excel to be exported.
''' </summary>
''' <remarks>
''' Created by Alex Liang on 4-Nov-2011
''' V3.0 upgrade by Alex Liang on 16-Nov-2011
''' V3.1 upgrade by Alex Liang on 14-Dec-2011 -- Replaced the Hashtable to Dictionary as the output sequence of Hashtable keys is inconsistent with the input
''' V3.2 upgrade by Alex Liang on 11-Jan-2012 -- Introduced more number formats; Optimized the XML skeleton.
''' V4.0 upgrade by Alex Liang on 27-Jan-2012 -- Switched to the brand new template-based style settings; Built-in multi-table supported.
''' V4.1 upgrade by Alex Liang on 27-Apr-2012 -- No change involved in Formatter.
''' V4.2 upgarde by Alex Liang on 3-May-2012  -- ReportName property excluded; Class name typo corrected.
''' V4.3 upgrade by Alex Liang on 19-Jul-2012 -- No change involved in Formatter.
''' V4.4 upgrade by Alex Liang on 30-Jul-2012 -- Template property introduced to accept template in XmlDocument type.
''' </remarks>

Public Class Formatter
    Private iMaxRow As Integer
    Private strPath As String
    Private xmlDoc As XmlDocument
    Private arlFlds As ArrayList
    Private dsSrc As DataSet

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
    ''' The template of XmlDocument type.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>It should be an instance of XmlDocument following OOXML format. If applied, the TemplatePath property would be omitted.</remarks>
    Public Property Template As XmlDocument
        Get
            Return xmlDoc
        End Get
        Set(ByVal value As XmlDocument)
            xmlDoc = value
        End Set
    End Property

    ''' <summary>
    ''' The path to the template.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The template should be an OOXML file.</remarks>
    Public Property TemplatePath As String
        Get
            Return strPath
        End Get
        Set(ByVal value As String)
            strPath = value
        End Set
    End Property

    ''' <summary>
    ''' The values of all the fields (like !Field_0).
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The ordinal and data type should exactly match those of the fields.</remarks>
    Public Property Fields As ArrayList
        Get
            Return arlFlds
        End Get
        Set(ByVal value As ArrayList)
            arlFlds = value
        End Set
    End Property

    ''' <summary>
    ''' Data source of the tables to be exported.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>It can be DataTable or DataSet.</remarks>
    Public Property DataSource As DataSet
        Get
            Return dsSrc
        End Get
        Set(ByVal value As DataSet)
            dsSrc = value
        End Set
    End Property

    ''' <summary>
    ''' Constructor.
    ''' </summary>
    ''' <remarks>Set the default values.</remarks>
    Public Sub New()
        iMaxRow = 65536
        arlFlds = New ArrayList
    End Sub
End Class