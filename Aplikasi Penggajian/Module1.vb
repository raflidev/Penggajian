Imports System.Data.SqlClient

Module Module1

    Public conn As SqlConnection
    Public da As SqlDataAdapter
    Public ds As DataSet
    Public cmd As SqlCommand
    Public dr As SqlDataReader

    Public Sub koneksi()
        'DESKTOP-D3GMB8G\SQLEXPRESS
        'data source = nama server
        'initial catalog = nama database
        'user id = sa; password = ??
        'integrated security = ?
        conn = New SqlConnection("data source=.\SQLEXPRESS;initial catalog=db12rpl3;integrated security=true")
        conn.Open()

    End Sub
End Module
