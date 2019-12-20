Imports System.Data.SqlClient

Public Class Pendapatan

    Sub kosongkan()
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

    Sub databaru()
        TextBox2.Clear()
    End Sub


    Sub Ketemu()
        TextBox1.Text = dr(0) 'id pendapatan
        TextBox2.Text = dr(1) 'nama pendapatan
    End Sub

    Sub tampilgrid()
        Call koneksi()
        da = New SqlDataAdapter("select * from TBLPendapatan", conn)
        ds = New DataSet
        da.Fill(ds)
        DGV.DataSource = ds.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Sub carikode()
        Call koneksi()
        cmd = New SqlCommand("select * from TBLpendapatan where id_pendapatan='" & TextBox1.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
    End Sub
    Private Sub Pendapatan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call tampilgrid()
        Me.CenterToScreen()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data belum lengkap")
        Else
            Call carikode()
            If dr.HasRows Then
                MsgBox("data kode sudah ada")
            Else
                Call koneksi()

                Dim simpan As String = "insert into TBLPENDAPATAN values('" & TextBox1.Text & "','" & TextBox2.Text & "')"

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
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("data masih kosong")
        Else
            Call carikode()
            If dr.HasRows Then
                Call koneksi()
                Dim update As String = "update tblPendapatan set nama_pendapatan='" & TextBox2.Text & "' where id_pendapatan='" & TextBox1.Text & "'"
                cmd = New SqlCommand(update, conn)
                dr = cmd.ExecuteReader
                MsgBox("data berhasil diperbaharui")
                Call kosongkan()
                Call tampilgrid()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("data masih kosong")
        Else
            If MessageBox.Show("yakin akan dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call carikode()
                If dr.HasRows Then
                    Call koneksi()
                    Dim delete As String = "delete from tblpendapatan where id_pendapatan='" & TextBox1.Text & "'"
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
        da = New SqlDataAdapter("select * from tblpendapatan where nama_pendapatan like '%" & TextBox5.Text & "%'", conn)
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