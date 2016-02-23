Imports System.Xml
Imports System.Configuration
Imports System.Text
Imports System.IO
Imports System.Security.Cryptography
Imports System.Net.Mail
Imports Emerson.Fisher.FLEx
Imports Emerson.Fisher.FLEx.Datalayer
Imports System.DirectoryServices

Public Class ErrorPage
    Inherits System.Web.UI.Page

    ''' <summary>
    ''' Load exception information from the temp XML file and fetch user name and Email account.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Initialize the page.</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim sbFilePath As New StringBuilder
            sbFilePath.Append(ConfigurationManager.AppSettings("TempPath"))
            sbFilePath.Append("/")
            sbFilePath.Append(HttpContext.Current.User.Identity.Name.Split("\")(1))
            sbFilePath.Append(".xml")

            Dim objXML As New XmlDocument
            objXML.Load(sbFilePath.ToString)
            GetUserInfo(objXML.GetElementsByTagName("UserID")(0).InnerText)
            txtApp.Text = objXML.GetElementsByTagName("AppName")(0).InnerText
            txtUserName.Text = hdnUserName.Value
            txtDate.Text = objXML.GetElementsByTagName("Time")(0).InnerText
            txtErrMsg.Text = objXML.GetElementsByTagName("ErrMsg")(0).InnerText
        End If
    End Sub

    ''' <summary>
    ''' Prepare the body of the Email to be sent.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Fire when the SendEmail is on click.</remarks>
    Protected Sub btnSendEmail_OnClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnSendEmail.Click
        'Below code is for email sending 
        Dim aInfo As String
        aInfo = txtAInfo.Text.ToString()
        If aInfo = "<Pleae enter the additional informaion to support team>" Then aInfo = ""

        Dim objReader As New StreamReader(Server.MapPath("~/Templates/email.htm"))
        Dim sbBody As New StringBuilder
        sbBody.Append(objReader.ReadToEnd())
        sbBody.Replace("##Application##", txtApp.Text)
        sbBody.Replace("##UserID##", txtUserName.Text)
        sbBody.Replace("##DateTime##", txtDate.Text)
        sbBody.Replace("##ErrorMessage##", txtErrMsg.Text)
        sbBody.Replace("##AdditionalInformation##", aInfo)

        SendEmail(ConfigurationManager.AppSettings("Sender"), ConfigurationManager.AppSettings("Receipt"), hdnUserMail.Value, sbBody.ToString, "FLEx - Error Page Demo")

        Response.Write("<script language=JavaScript>alert('Your service request has been sent to IT HelpDesk');</script>")
    End Sub

    ''' <summary>
    ''' Send Email
    ''' </summary>
    ''' <param name="SenderEmailAddress">sender address</param>
    ''' <param name="RecipientEmailAddress">receive address</param>
    ''' <param name="CC">cc address</param>
    ''' <param name="MessageBody">message content</param>
    ''' <param name="Subject">subject</param>
    ''' <remarks></remarks>
    Sub SendEmail(ByVal SenderEmailAddress As String, ByVal RecipientEmailAddress As String, ByVal CC As String, ByVal MessageBody As String, ByVal Subject As String)
        Try
            Dim mMailMessage As New MailMessage()
            mMailMessage.From = New MailAddress(SenderEmailAddress)
            Dim SplitStr1() As String = Split(RecipientEmailAddress, ";")
            Dim i As Integer
            For i = 0 To UBound(SplitStr1)
                mMailMessage.To.Add(New MailAddress(SplitStr1(i)))

            Next
            If Not CC Is Nothing And CC <> String.Empty Then
                Dim SplitStr2() As String = Split(CC, ";")
                For i = 0 To UBound(SplitStr2)
                    mMailMessage.CC.Add(New MailAddress(SplitStr2(i)))
                Next
            End If

            mMailMessage.Subject = Subject
            mMailMessage.Body = MessageBody
            mMailMessage.IsBodyHtml = True
            mMailMessage.Priority = MailPriority.Normal
            Dim mSmtpClient As New SmtpClient()
            mSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
            mSmtpClient.Host = "inetmail.emrsn.net"
            mSmtpClient.Port = 25
            mSmtpClient.Send(mMailMessage)
            lblmsg.Visible = True
        Catch ex As Exception
            'Throw ex
            'Dim ErrorMsg As String = ex.GetBaseException.ToString
            Console.WriteLine("Email:Exception occured when emailing" + ex.GetBaseException.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Get the name and Email address of the user who raises the exception.
    ''' </summary>
    ''' <param name="strUserID">The Windows login ID of the user in question. It may be prefixed with domain ID.</param>
    ''' <remarks>Store the values in hidden fields.</remarks>
    Private Sub GetUserInfo(ByVal strUserID As String)
        Dim objSearchRoot As New DirectoryEntry(ConfigurationManager.AppSettings("ADRoot"))
        Dim objDSUserPath As New DirectorySearcher(objSearchRoot)

        Dim sbFilter As New StringBuilder
        sbFilter.Append("(sAMAccountName=")
        sbFilter.Append(strUserID.Split("\")(1))   ' Cut the domain ID if it presents.
        sbFilter.Append(")")

        objDSUserPath.Filter = sbFilter.ToString
        Dim objResult As SearchResult = objDSUserPath.FindOne

        hdnUserName.Value = objResult.Properties("cn").Item(0)
        hdnUserMail.Value = objResult.Properties("mail").Item(0)
    End Sub
End Class