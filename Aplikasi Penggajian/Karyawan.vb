Imports System.Data.SqlClient

Public Class Karyawan
    Sub notis()
        Dim tahun As String = Year(Today)
        Call koneksi()
        cmd = New SqlCommand("select id_karyawan from tblkaryawan order by 1 desc", conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If Not dr.HasRows Then
            tidKaryawan.Text = tahun + "000001"
        Else
            tidKaryawan.Text = Format(Microsoft.VisualBasic.Right(dr.Item("id_karyawan"), 10) + 1, "0000000000")
        End If
    End Sub
    Sub carikode()
        Call koneksi()
        cmd = New SqlCommand("select * from tblkaryawan where id_karyawan='" & tidKaryawan.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
    End Sub
    Sub Ketemu()
        tidKaryawan.Text = dr(0)
        namaKaryawan.Text = dr(1)
        tjabatan.Text = dr(2)
        tGajiPokok.Text = dr(3)
        tTunjanganJabatan.Text = dr(4)
        tTunjanganKeluarga.Text = dr(5)
        tTunjunganAnak.Text = dr(6)
        tUangLembur.Text = dr(7)
        tUangMakan.Text = dr(8)
        tUangTransport.Text = dr(9)
        tStatusPernikahan.Text = dr(10)
        tJumlahAnak.Text = dr(11)
    End Sub
    Sub jabatan()
        Call koneksi()
        cmd = New SqlCommand("select distinct jabatan from tblkaryawan", conn)
        dr = cmd.ExecuteReader
        tjabatan.Items.Clear()
        Do While dr.Read
            tjabatan.Items.Add(dr(0))
        Loop
    End Sub
    Sub statusmenikah()
        tStatusPernikahan.Items.Clear()
        tStatusPernikahan.Items.Add("MENIKAH")
        tStatusPernikahan.Items.Add("BELUM MENIKAH")
    End Sub

    Sub bersihkan()
        tidKaryawan.Clear()
        namaKaryawan.Clear()
        tjabatan.Text = ""
        tGajiPokok.Clear()
        tTunjanganJabatan.Clear()
        tTunjanganKeluarga.Clear()
        tTunjunganAnak.Clear()
        tUangLembur.Clear()
        tUangMakan.Clear()
        tUangTransport.Clear()
        tStatusPernikahan.Text = ""
        tJumlahAnak.Clear()
    End Sub

    Sub tampilgrid()
        Call koneksi()
        da = New SqlDataAdapter("select * from tblkaryawan", conn)
        ds = New DataSet
        da.Fill(ds)
        DGV.DataSource = ds.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox13.TextChanged
        Call koneksi()
        da = New SqlDataAdapter("select * from tblkaryawan where nama_karyawan like '%" & TextBox13.Text & "%'", conn)
        ds = New DataSet
        da.Fill(ds)
        DGV.DataSource = ds.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Private Sub Karyawan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        Call notis()
        Call jabatan()
        Call tampilgrid()
        Call statusmenikah()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call koneksi()
        cmd = New SqlCommand("delete from tblkaryawan where id_karyawan='" & tidKaryawan.Text & "'", conn)
        dr = cmd.ExecuteReader
        MsgBox("data berhasil dihapus")
        Call bersihkan()
        Call tampilgrid()


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call bersihkan()
        Call notis()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If tidKaryawan.Text = "" Or namaKaryawan.Text = "" Or tjabatan.Text = "" Or tGajiPokok.Text = "" Or tTunjanganJabatan.Text = "" Or tTunjanganKeluarga.Text = "" Or tTunjunganAnak.Text = "" Or tUangLembur.Text = "" Or tUangMakan.Text = "" Or tUangTransport.Text = "" Or tStatusPernikahan.Text = "" Or tJumlahAnak.Text = "" Then
            MsgBox("data masih kosong")
        End If


        Call koneksi()
        cmd = New SqlCommand("insert into tblkaryawan values('" & tidKaryawan.Text & "','" & namaKaryawan.Text & "','" & tjabatan.Text & "'," & tGajiPokok.Text & "," & tTunjanganJabatan.Text & "," & tTunjanganKeluarga.Text & "," & tTunjunganAnak.Text & "," & tUangLembur.Text & "," & tUangMakan.Text & "," & tUangTransport.Text & ",'" & tStatusPernikahan.Text & "'," & tJumlahAnak.Text & ")", conn)
        dr = cmd.ExecuteReader
        MsgBox("data berhasil ditambahkan")
        Call bersihkan()
        Call tampilgrid()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If tidKaryawan.Text = "" Or namaKaryawan.Text = "" Or tjabatan.Text = "" Or tGajiPokok.Text = "" Or tTunjanganJabatan.Text = "" Or tTunjanganKeluarga.Text = "" Or tTunjunganAnak.Text = "" Or tUangLembur.Text = "" Or tUangMakan.Text = "" Or tUangTransport.Text = "" Or tStatusPernikahan.Text = "" Or tJumlahAnak.Text = "" Then
            MsgBox("data masih kosong")
        End If
        Call koneksi()
        cmd = New SqlCommand("update tblkaryawan set nama_karyawan='" & namaKaryawan.Text & "',jabatan='" & tjabatan.Text & "',gaji_pokok=" & tGajiPokok.Text & ",tunjangan_jabatan=" & tTunjanganJabatan.Text & ",tunjangan_keluarga=" & tTunjanganKeluarga.Text & ",tunjangan_anak=" & tTunjunganAnak.Text & ",uang_lembur=" & tUangLembur.Text & ",uang_makan=" & tUangMakan.Text & ",uang_transport=" & tUangTransport.Text & ",status_pernikahan='" & tStatusPernikahan.Text & "',jumlah_anak=" & tJumlahAnak.Text & " where id_karyawan='" & tidKaryawan.Text & "'", conn)
        dr = cmd.ExecuteReader
        MsgBox("data berhasil diperbaharui")
        Call bersihkan()
        Call tampilgrid()
    End Sub

    Private Sub DGV_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellContentClick

    End Sub

    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        tidKaryawan.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        Call carikode()
        If dr.HasRows Then
            Call Ketemu()

        End If
    End Sub
End Class