Public Class Data_Access_Class

    Private vDB As New DB_SQLServer

    Public Function Get_Customers() As DataSet

        Dim v_DataSet As DataSet

        v_DataSet = vDB.ReturnDataSet("EXEC usp_Customers_Select")
        vDB.CloseDB()

        Return v_DataSet
    End Function

End Class
