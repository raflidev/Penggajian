Imports System.Data.SqlClient

Public Class Login
    Sub bersihkan()
        tUsername.Clear()
        tPassword.Clear()

    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If tUsername.Text = "" Or tPassword.Text = "" Then
            MsgBox("username/password salah")
        Else
            Call koneksi()
            cmd = New SqlCommand("select * from tbluser where nama_user='" & tUsername.Text & "' and pwd_user='" & tPassword.Text & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                Me.Visible = False
                MenuUtama.panelkode.Text = dr(0)
                MenuUtama.paneluser.Text = dr(1)
                MenuUtama.panelstatus.Text = dr(2)
                MenuUtama.Show()

            Else
                MsgBox("username/password salah")
            End If
        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Call bersihkan()

    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()


    End Sub
End Class
