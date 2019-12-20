Imports System.Data.SqlClient

Public Class Penggajian
    Sub bersihkan()
        Tidgaji.Text = ""
        Tidkaryawan.Text = ""
        Tnamakaryawan.Text = ""
        Tstatus.Text = ""
        Tsakit.Text = ""
        Tjumlahanak.Text = ""
        Tjabatan.Text = ""
        Tketerangan.Text = ""
        TTotalpendapatan.Text = 0
        TTotalpotongan.Text = 0
        Tgajibersih.Text = 0
        tjumlahharimasuk.Text = 0
        Tjumlahlembur.Text = 0
        DGV1.Columns(0).ReadOnly = True
        DGV1.Columns(1).ReadOnly = True
        For baris1 As Integer = 0 To DGV1.RowCount - 1
            DGV1.Rows(baris1).Cells(2).Value = 0
            DGV1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Next

        DGV2.Columns(0).ReadOnly = True
        DGV2.Columns(1).ReadOnly = True
        For baris2 As Integer = 0 To DGV2.RowCount - 1
            DGV2.Rows(baris2).Cells(2).Value = 0
            DGV2.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Next

        Call notis()    
    End Sub


    Sub notis()

        Call koneksi()
        cmd = New SqlCommand("select id_gaji from tblgaji order by 1 desc", conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If Not dr.HasRows Then
            Tidgaji.Text = "0000000001"
        Else
            Tidgaji.Text = Format(Microsoft.VisualBasic.Right(dr.Item("id_gaji"), 10) + 1, "0000000000")
        End If
    End Sub
    Sub tampilpendapatan()
        DGV1.Columns.Clear()

        Call koneksi()
        da = New SqlDataAdapter("select * from tblpendapatan", conn)
        ds = New DataSet
        da.Fill(ds)
        DGV1.DataSource = ds.Tables(0)
        DGV1.Columns.Add("jumlah", "Jumlah")
        DGV1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

        For baris As Integer = 0 To DGV1.RowCount - 2
            DGV1.Rows(baris).Cells(2).Value = 0
        Next

        DGV1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DGV1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

    End Sub
    Sub tampilpotongan()
        DGV2.Columns.Clear()
        Call koneksi()

        da = New SqlDataAdapter("select * from tblpotongan", conn)
        ds = New DataSet
        da.Fill(ds)
        DGV2.DataSource = ds.Tables(0)

        DGV2.Columns.Add("jumlah", "Jumlah Potongan")
        DGV2.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

        For baris As Integer = 0 To DGV2.RowCount - 2
            DGV2.Rows(baris).Cells(2).Value = 0
        Next

        DGV2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DGV2.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
    End Sub

    Sub cariTotalPendapatan()
        Try
            Dim hitung As Integer = 0
            For baris As Integer = 0 To DGV1.RowCount - 1
                hitung = hitung + DGV1.Rows(baris).Cells(2).Value
                TTotalpendapatan.Text = hitung

            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub caritotalpotongan()
        Try
            Dim hitung As Integer = 0
            For baris As Integer = 0 To DGV2.RowCount - 1
                hitung = hitung + DGV2.Rows(baris).Cells(2).Value
                TTotalpotongan.Text = hitung
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Tentri_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tentri.Click

    End Sub

    Private Sub Penggajian_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        tanggalawal.Text = DateAdd(DateInterval.Month, -1, tanggalakhir.Value)
        tanggalakhir.Text = Today

        Tjumlahhari.Text = DateDiff(DateInterval.Day, tanggalawal.Value, tanggalakhir.Value) - 4

    End Sub

    Private Sub Penggajian_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        Call notis()
        Call cariTotalPendapatan()
        Call caritotalpotongan()
        Tentri.Text = Today
        Call tampilpotongan()
        Call tampilpendapatan()
        Call koneksi()
        cmd = New SqlCommand("select * from tblkaryawan", conn)
        dr = cmd.ExecuteReader
        Do While dr.Read
            Tidkaryawan.Items.Add(dr(0))
        Loop
        Call koneksi()
        cmd = New SqlCommand("select * from tblkaryawan", conn)
        dr = cmd.ExecuteReader
        Do While dr.Read
            Tnamakaryawan.Items.Add(dr(1))
        Loop
    End Sub

    Private Sub Tidkaryawan_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Tidkaryawan.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            cmd = New SqlCommand("select id_karyawan from tblgaji where id_karyawan='" & Tidkaryawan.Text & "' and month(tanggal)='" & Month(Tentri.Text) & "' and year(tanggal) = '" & Year(Tentri.Text) & "'", conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                Call bersihkan()
                MsgBox("data gaji karyawan tersebut bulan dan tahun ini sudah dientri")
                Exit Sub
            Else
                Call koneksi()
                cmd = New SqlCommand("select * from tblkaryawan where id_karyawan='" & Tidkaryawan.Text & "'", conn)
                dr = cmd.ExecuteReader
                dr.Read()
                If Not dr.HasRows Then
                    Call bersihkan()
                    MsgBox("ID karyawan tidak terdaftar")
                    Exit Sub
                Else
                    tjumlahharimasuk.Text = Tjumlahhari.Text
                    Tnamakaryawan.Text = dr.Item("nama_karyawan")
                    Tstatus.Text = dr.Item("status_pernikahan")
                    Tjumlahanak.Text = dr.Item("jumlah_anak")
                    Tjabatan.Text = dr.Item("jabatan")
                    DGV1.Rows(0).Cells(2).Value = dr.Item("gaji_pokok")
                    DGV1.Rows(1).Cells(2).Value = dr.Item("tunjangan_jabatan")

                    If Tstatus.Text = "BELUM MENIKAH" Then
                        DGV1.Rows(2).Cells(2).Value = 0
                    Else
                        DGV1.Rows(2).Cells(2).Value = dr.Item("tunjangan_keluarga")
                    End If

                    If Tjumlahanak.Text > 3 Then
                        DGV1.Rows(3).Cells(2).Value = 3 * dr.Item("tunjangan_anak")
                    Else
                        DGV1.Rows(3).Cells(2).Value = Tjumlahanak.Text * dr.Item("tunjangan_anak")
                    End If

                    DGV1.Rows(4).Cells(2).Value = dr.Item("uang_transport") * tjumlahharimasuk.Text
                    DGV1.Rows(5).Cells(2).Value = dr.Item("uang_makan") * tjumlahharimasuk.Text
                    DGV1.Rows(6).Cells(2).Value = dr.Item("uang_lembur") * Tjumlahlembur.Text

                    DGV1.Rows(0).ReadOnly = True
                    DGV1.Rows(1).ReadOnly = True
                    DGV1.Rows(2).ReadOnly = True
                    DGV1.Rows(3).ReadOnly = True
                    DGV1.Rows(4).ReadOnly = True
                    DGV1.Rows(5).ReadOnly = True
                    DGV1.Rows(6).ReadOnly = True
                    Call cariTotalPendapatan()

                    DGV2.Rows(1).Cells(2).Value = DGV1.Rows(0).Cells(2).Value * 1 / 100
                    DGV2.Rows(0).Cells(2).Value = DGV1.Rows(0).Cells(2).Value * 4 / 100
                    Call caritotalpotongan()
                End If
            End If

        End If
    End Sub

    Sub hitungmasukkerja()
        On Error Resume Next
        tjumlahharimasuk.Text = Val(Tjumlahhari.Text) - (Val(Tsakit.Text) + Val(Tizin.Text) + Val(Talpa.Text))
        tjumlahharimasuk.Enabled = False

    End Sub

    Private Sub Tidkaryawan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tidkaryawan.SelectedIndexChanged
        Call koneksi()
        cmd = New SqlCommand("select * from tblkaryawan where id_karyawan = " & Tidkaryawan.Text & "", conn)

        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Tnamakaryawan.Text = dr("nama_karyawan")
            Tstatus.Text = dr("status_pernikahan")
            Tjabatan.Text = dr("jabatan")
            Tjumlahanak.Text = dr("jumlah_anak")
            '===============
            DGV1.Rows(0).Cells(2).Value = dr("gaji_pokok")
            DGV1.Rows(1).Cells(2).Value = dr("tunjangan_jabatan")
            DGV1.Rows(2).Cells(2).Value = dr("tunjangan_keluarga")
            DGV1.Rows(3).Cells(2).Value = dr("tunjangan_anak")
            DGV1.Rows(4).Cells(2).Value = dr("uang_lembur")
            DGV1.Rows(5).Cells(2).Value = dr("uang_makan")
            DGV1.Rows(6).Cells(2).Value = dr("uang_transport")
        End If
    End Sub



    Private Sub DGV1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV1.CellContentClick

    End Sub


    Private Sub DGV1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV1.CellEndEdit
        If e.ColumnIndex = 2 Then
            Call cariTotalPendapatan()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call bersihkan()
    End Sub

    Private Sub DGV2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV2.CellContentClick

    End Sub

    Private Sub DGV2_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV2.CellEndEdit
        If e.ColumnIndex = 2 Then
            Call caritotalpotongan()
        End If
    End Sub

    Private Sub TTotalpendapatan_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TTotalpendapatan.TextChanged
        On Error Resume Next
        Tgajibersih.Text = Val(TTotalpendapatan.Text) - Val(TTotalpotongan.Text)
    End Sub

    Private Sub TTotalpotongan_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TTotalpotongan.TextChanged
        On Error Resume Next
        Tgajibersih.Text = Val(TTotalpendapatan.Text) - Val(TTotalpotongan.Text)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        cmd = New SqlCommand("insert into tblgaji values('" & Tidgaji.Text & "','" & Format(DateValue(Tentri.Text), "MM/dd/yyyy") & "', '" & Tidkaryawan.Text & "','" & tanggalawal.Text & "','" & tanggalakhir.Text & "','" & Tjumlahhari.Text & "','" & Tsakit.Text & "','" & Tizin.Text & "','" & Talpa.Text & "','" & tjumlahharimasuk.Text & "','" & TTotalpendapatan.Text & "','" & TTotalpotongan.Text & "','" & Tgajibersih.Text & "','" & Tketerangan.Text & "','USR01')", conn)
        cmd.ExecuteNonQuery()

        For baris As Integer = 0 To DGV1.RowCount - 2
            Call koneksi()
            cmd = New SqlCommand("insert into tblgajidetail values('" & Tidgaji.Text & "','" & DGV1.Rows(baris).Cells(0).Value & "','" & DGV1.Rows(baris).Cells(2).Value & "','-',0)", conn)
            cmd.ExecuteNonQuery()
        Next
        'update
        For baris As Integer = 0 To DGV2.RowCount - 2
            Call koneksi()
            cmd = New SqlCommand("select id_gaji,id_pendapatan from tblgajidetail where id_gaji = '" & Tidgaji.Text & "' and id_pendapatan='" & DGV1.Rows(baris).Cells(0).Value & "'", conn)
            dr = cmd.ExecuteReader()
            dr.Read()
            If dr.HasRows Then
                Call koneksi()
                cmd = New SqlCommand("update tblgajidetail set id_potongan='" & DGV2.Rows(baris).Cells(0).Value & "',jumlah_potongan=" & DGV2.Rows(baris).Cells(2).Value & " where id_gaji='" & dr(0) & "' and id_pendapatan='" & dr(1) & "'", conn)
                cmd.ExecuteNonQuery()
            End If
        Next

        If MessageBox.Show("cetak slip gaji....?", "Perhatian", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Cetak.show()
            Cetak.CRV.ReportSource = "Slipgaji1.rpt"
            Cetak.CRV.RefreshReport()
        End If
        Call bersihkan()

    End Sub

    Private Sub Tsakit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tsakit.TextChanged
        Call hitungmasukkerja()

    End Sub

    Private Sub Tizin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tizin.TextChanged
        Call hitungmasukkerja()

    End Sub

    Private Sub Talpa_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Talpa.TextChanged
        Call hitungmasukkerja()

    End Sub

    Private Sub Tnamakaryawan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tnamakaryawan.SelectedIndexChanged

    End Sub

    Private Sub Tjumlahlembur_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Tjumlahlembur.TextChanged
        On Error Resume Next
        Call koneksi()

        cmd = New SqlCommand("select * from tblkaryawan where id_karyawan='" & Tidkaryawan.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            DGV1.Rows(6).Cells(2).Value = dr.Item("uang_lembur") * Tjumlahlembur.Text
        Else
            DGV1.Rows(6).Cells(2).Value = 0
        End If
        Call cariTotalPendapatan()
        Call caritotalpotongan()

    End Sub

    Private Sub tjumlahharimasuk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tjumlahharimasuk.Click

    End Sub

    Private Sub tjumlahharimasuk_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tjumlahharimasuk.TextChanged
        On Error Resume Next
        DGV1.Rows(4).Cells(2).Value = dr.Item("uang_transport") * tjumlahharimasuk.Text
        DGV1.Rows(5).Cells(2).Value = dr.Item("uang_makan") * tjumlahharimasuk.Text
        DGV1.Rows(6).Cells(2).Value = dr.Item("uang_lembur") * tjumlahharimasuk.Text
        Call cariTotalPendapatan()
        Call caritotalpotongan()

    End Sub
End Class
