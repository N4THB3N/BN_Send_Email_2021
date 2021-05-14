Imports System.Data.SqlClient
Imports System.Configuration
Public Class DB_SQLServer

    Private v_SQLConnection As SqlConnection
    Private v_SQLCommand As SqlCommand
    Private v_DataReader As IDataReader

    Private Function OpenDB() As Boolean

        v_SQLConnection = New SqlConnection

        Try

            If v_SQLConnection.State = ConnectionState.Open Then
                v_SQLConnection.Close()
            End If

            v_SQLConnection.ConnectionString = ConfigurationManager.ConnectionStrings("BN").ToString
            v_SQLConnection.Open()
            OpenDB = True

        Catch ex As Exception
            OpenDB = False
        End Try

    End Function

    Public Function CloseDB() As Boolean

        Try

            ' Checks if there is an open data reader to the database
            If Not v_DataReader Is Nothing Then
                If Not v_DataReader.IsClosed Then
                    v_DataReader.Close()
                End If
            End If

            ' Closes the connection
            If v_SQLConnection.State = ConnectionState.Open Then
                v_SQLConnection.Close()
            End If

            CloseDB = True
        Catch ex As Exception
            CloseDB = False
        End Try

    End Function

    ' Returns a DataSet
    Public Function ReturnDataSet(ByVal pQuery As String) As DataSet
        v_SQLCommand = New SqlCommand
        Dim vDataSet As New DataSet
        Dim vDataAdapter As New SqlDataAdapter

        If OpenDB() Then
            v_SQLCommand.Connection = v_SQLConnection
            v_SQLCommand.CommandType = CommandType.Text
            v_SQLCommand.CommandText = pQuery
            v_SQLCommand.CommandTimeout = 3600
            vDataAdapter = New SqlDataAdapter(pQuery, v_SQLConnection)
            vDataAdapter.Fill(vDataSet)
            CloseDB()
            Return vDataSet
        End If
        Return Nothing
    End Function

End Class
