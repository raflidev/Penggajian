Imports System.Data.SqlClient
Public Class Cetak

    Private Sub Cetak_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()

        cmd = New SqlCommand("select nama_karyawan from tblkaryawan", conn)
        dr = cmd.ExecuteReader

        tidKaryawan.Items.Clear()
        Do While dr.Read()
            tidKaryawan.Items.Add(dr(0))
        Loop
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CRV.SelectionFormula = "{tblkaryawan.nama_karyawan} = '" & tidKaryawan.Text & "'"
        CRV.ReportSource = "Slipgaji1.rpt"
        CRV.RefreshReport()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        CRV.ReportSource = "LaporanKaryawan.rpt"
        CRV.RefreshReport()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CRV.ReportSource = "LaporanPotongan.rpt"
        CRV.RefreshReport()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        CRV.ReportSource = "LaporanPendapatan.rpt"
        CRV.RefreshReport()
    End Sub

    Private Sub GroupBox5_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox5.Enter

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CRV.SelectionFormula = "{tblgajidetail.id_potongan} = '02'"
        CRV.ReportSource = "LaporanPotonganBPJS.rpt"
        CRV.RefreshReport()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        CRV.SelectionFormula = "{tblgajidetail.id_potongan} = '01'"
        CRV.ReportSource = "LaporanPotonganPPH21.rpt"
        CRV.RefreshReport()
    End Sub
End Class