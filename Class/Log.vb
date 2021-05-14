Imports System.IO

Public Class Log

    Public Shared Sub Append_To_Log(ByVal p_LogContent As String)

        Dim v_Path As String = AppDomain.CurrentDomain.BaseDirectory & "\Logs\" & Now.Date.ToString("yyyyMM") & "\"

        Dim v_FileName As String = "BN_MailJet_Sender_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"

        ' If directory does not exists (first day of month), then create it.
        If Not IO.Directory.Exists(v_Path) Then

            IO.Directory.CreateDirectory(v_Path)
        End If

        Dim v_FileStream As FileStream = File.Open(v_Path & v_FileName, FileMode.Append)
        Dim v_StreamWriter As New StreamWriter(v_FileStream)

        v_StreamWriter.WriteLine(String.Format("{1}  {0} ", p_LogContent, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")))
        v_StreamWriter.Flush()
        v_StreamWriter.Close()
        v_FileStream.Close()

    End Sub

End Class
