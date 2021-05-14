Imports System.Configuration
Imports System.Net
Imports System.Net.Mail

Public Class Email_Sender

    Public Function SendEmail_ThroughMailJet(ByVal p_To As String, ByVal p_Cc As String, ByVal p_Bcc As String, ByVal p_Reply_To As String, ByVal p_Subject As String, ByVal p_Body As String) As String
        Dim v_Result As String = ""

        Try
            Dim v_From As String
            v_From = ConfigurationManager.AppSettings("SendEmailFrom")

            Dim v_SMTP As New SmtpClient()
            v_SMTP.Host = "in-v3.mailjet.com"
            v_SMTP.EnableSsl = True
            v_SMTP.UseDefaultCredentials = True

            v_SMTP.Credentials = New NetworkCredential(ConfigurationManager.AppSettings("MailJet_KeyID"), ConfigurationManager.AppSettings("MailJet_SecretKey"))
            v_SMTP.Port = Integer.Parse(ConfigurationManager.AppSettings("MailJet_Port"))

            Using v_Mail_Message As New MailMessage(v_From, p_To)
                If p_Cc.ToString.Length > 0 Then
                    Dim v_Cc As MailAddress = New MailAddress(p_Cc)
                    v_Mail_Message.CC.Add(v_Cc)
                End If

                If p_Bcc.ToString.Length > 0 Then
                    Dim v_Bcc_List As String()

                    v_Bcc_List = p_Bcc.Split(",")

                    For Each Item In v_Bcc_List
                        Dim v_Bcc As MailAddress = New MailAddress(Item)
                        v_Mail_Message.Bcc.Add(Item)
                    Next

                End If

                If p_Reply_To.ToString.Length > 0 Then
                    Dim v_Reply_To As MailAddress = New MailAddress(p_Reply_To)
                    v_Mail_Message.ReplyToList.Add(v_Reply_To)
                End If

                v_Mail_Message.Subject = p_Subject
                v_Mail_Message.Body = p_Body
                v_Mail_Message.IsBodyHtml = True

                v_SMTP.Send(v_Mail_Message)


            End Using

            v_Result = "OK"


        Catch ex As Exception
            v_Result = ex.Message.ToString()
        End Try

        Return v_Result

    End Function

End Class
