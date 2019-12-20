Imports System.Data.SqlClient

Public Class User

    Sub kosongkan()
        TextBox1.Clear()
        TextBox2.Clear()
        ComboBox1.Text = ""
        TextBox3.Clear()
    End Sub

    Sub databaru()
        TextBox2.Clear()
        ComboBox1.Text = ""
        TextBox3.Clear()
    End Sub

    Sub Ketemu()
        TextBox2.Text = dr(1) 'nama user
        ComboBox1.Text = dr(2) 'status
        TextBox3.Text = dr(3) 'pass

    End Sub

    Sub tampilstatus()
        Call koneksi()
        cmd = New SqlCommand("select distinct status_user from tbluser", conn)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox1.Items.Add(dr(0))
        Loop
    End Sub

    Sub tampilgrid()
        Call koneksi()
        da = New SqlDataAdapter("select * from TBLUser", conn)
        ds = New DataSet
        da.Fill(ds)
        DGV.DataSource = ds.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Sub carikode()
        Call koneksi()
        cmd = New SqlCommand("select * from tbluser where id_user='" & TextBox1.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
    End Sub

    Private Sub User_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        Call tampilstatus()
        Call tampilgrid()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call carikode()
            If dr.HasRows Then
                Call Ketemu()
            Else
                Call databaru()
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Data belum lengkap")
        Else
            Call carikode()
            If dr.HasRows Then
                MsgBox("data kode sudah ada")
            Else
                Call koneksi()

                Dim simpan As String = "insert into tbluser values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox3.Text & "')"

                cmd = New SqlCommand(simpan, conn)
                dr = cmd.ExecuteReader
                MsgBox("data berhasil ditambah")
                Call kosongkan()
                Call tampilgrid()
            End If
        End If
           
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call kosongkan()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Then
            MsgBox("data masih kosong")
        Else
            Call carikode()
            If dr.HasRows Then
                Call koneksi()
                Dim update As String = "update tbluser set nama_user='" & TextBox2.Text & "',Status_user='" & ComboBox1.Text & "',pwd_user='" & TextBox3.Text & "' where id_user='" & TextBox1.Text & "'"
                cmd = New SqlCommand(update, conn)
                dr = cmd.ExecuteReader
                MsgBox("data berhasil diperbaharui")
                Call kosongkan()
                Call tampilgrid()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Then
            MsgBox("data masih kosong")
        Else
            If MessageBox.Show("yakin akan dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call carikode()
                If dr.HasRows Then
                    Call koneksi()
                    Dim delete As String = "delete from tbluser where id_user='" & TextBox1.Text & "'"
                    cmd = New SqlCommand(delete, conn)
                    dr = cmd.ExecuteReader
                    MsgBox("data berhasil dihapus")
                    Call kosongkan()
                    Call tampilgrid()
                End If
            Else
                Call kosongkan()
            End If
        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        da = New SqlDataAdapter("select * from tbluser where nama_user like '%" & TextBox5.Text & "%'", conn)
        ds = New DataSet
        da.Fill(ds)
        DGV.DataSource = ds.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Private Sub DGV_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellContentClick

    End Sub

    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        Call carikode()
        If dr.HasRows Then
            Call Ketemu()

        End If
    End Sub
End Class
