Imports System.Text
Imports System.Web
Imports System.Xml
Imports System.Diagnostics
Imports System.Configuration

''' <summary>
''' This class helps to log exception message into Windows event log.
''' </summary>
''' <remarks>
''' Created by Alex Liang on 13-Oct-2011
''' </remarks>
'''
Public Class ExceptionLogger
    Implements IHttpModule

    Private strEvtSrc As String
    Private strLogNm As String
    Private strRdctTo As String
    ''' <summary>
    ''' Log the exception to the Windows event log.
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="args"></param>
    ''' <remarks></remarks>
    Public Sub OnError(ByVal obj As Object, ByVal args As EventArgs)
        Dim lastError As Exception = HttpContext.Current.Server.GetLastError().GetBaseException()
        Dim sbMsg As New StringBuilder

        sbMsg.Append("Error Message:")
        sbMsg.Append(Environment.NewLine)
        sbMsg.Append(lastError.Message)
        sbMsg.Append(Environment.NewLine)
        sbMsg.Append(Environment.NewLine)
        sbMsg.Append("Stack Trace:")
        sbMsg.Append(Environment.NewLine)
        sbMsg.Append(lastError.StackTrace)
        WriteEntry(strEvtSrc, sbMsg.ToString(), EventLogEntryType.Error)

        HttpContext.Current.Response.Redirect(strRdctTo)
    End Sub
    ''' <summary>
    ''' This method will check the log folder exists in Windows Event viewer. If not exists this method will create the log folder and log the exception details in it.
    ''' </summary>
    ''' <param name="eventSrc">Event Source which will be displayed in Eventviewer</param>
    ''' <param name="exceptionMessage">Exception message</param>
    ''' <param name="logType">Type of the exception like error,warning,information etc.</param>
    ''' <remarks>To write the exception in</remarks>
    Private Sub WriteEntry(ByVal eventSrc As String, ByVal exceptionMessage As String, ByVal logType As EventLogEntryType)
        If Not EventLog.Exists(strLogNm) Or Not EventLog.SourceExists(eventSrc) Then
            EventLog.CreateEventSource(eventSrc, strLogNm)
        End If
        EventLog.WriteEntry(eventSrc, exceptionMessage, logType)
    End Sub

    ''' <summary>
    ''' Initiate the HttpModule
    ''' </summary>
    ''' <param name="objApp">Instance of the HttpApplication</param>
    ''' <remarks>To flesh out the Init method required by the IHttpModule interface. The system error </remarks>
    Public Sub Init(ByVal objApp As System.Web.HttpApplication) Implements System.Web.IHttpModule.Init
        strEvtSrc = ConfigurationManager.AppSettings("EventSource")
        strLogNm = ConfigurationManager.AppSettings("LogName")
        strRdctTo = ConfigurationManager.AppSettings("RedirectTo")
        AddHandler objApp.Error, AddressOf OnError
    End Sub

    ''' <summary>
    ''' Dispose the HttpModule
    ''' </summary>
    ''' <remarks>Nothing to do in this case.</remarks>
    Public Sub Dispose() Implements System.Web.IHttpModule.Dispose

    End Sub
End Class
