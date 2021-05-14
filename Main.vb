Imports BC_IFX_Mailjet_Sender_2021.Email_Sender
Imports BC_IFX_Mailjet_Sender_2021.Log
Imports BC_IFX_Mailjet_Sender_2021.Data_Access_Class
Module Main

    Sub Main()
        Try
            Dim v_Email_Sender As New Email_Sender
            Dim v_Result As String = ""
            Dim v_Customer_Info As DataSet
            Dim v_Data_Access As New Data_Access_Class

            Dim v_To As String = ""
            Dim v_Customer_Nm As String = ""
            Dim v_Customer_Last_Nm As String = ""
            Dim v_Body As String = ""

            Append_To_Log(" ************************************************** ")
            Append_To_Log(" *********   Mail Sender - Task Started    ******** ")
            Append_To_Log(" ... ")

            Append_To_Log(" ... Reading Customers ")

            v_Customer_Info = v_Data_Access.Get_Customers()

            For Each v_Cust In v_Customer_Info.Tables(0).Rows

                v_To = v_Cust("Email").ToString()
                v_Customer_Nm = v_Cust("Customer_Nm").ToString()
                v_Customer_Last_Nm = v_Cust("Customer_Last_Nm").ToString()

                v_Body = "Bienvenido nuevo usuario.


                          Es un gusto para nosotros atenderle.


                          Saludos."

                Append_To_Log("...... Row # : ")
                Append_To_Log("...... Customer_Nm: " & v_Customer_Nm & v_Customer_Last_Nm)
                Append_To_Log("...... Customer_Mail: " & v_To)

                v_Result = v_Email_Sender.SendEmail_ThroughMailJet(v_To, "", "", "", "Bienvenido a Janiza", v_Body)

                If v_Result = "OK" Then
                    Append_To_Log("........... Email Enviado.")
                Else
                    Append_To_Log("................ ERROR: " & v_Result)
                End If

                Append_To_Log("......")
                Append_To_Log("********** END **********")

            Next

        Catch ex As Exception
            Append_To_Log("................ ERROR: " & ex.InnerException.ToString())
            Append_To_Log("......")
            Append_To_Log("********** END **********")
        End Try
    End Sub

End Module
