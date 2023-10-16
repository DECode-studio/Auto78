Imports System.Runtime.InteropServices
Imports MySql.Data.MySqlClient
Imports System.Data.OleDb

Public Class frm_main

    Dim Pos As Point

    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As Int32
    End Function

    Private Sub frm_main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Animasi button menu
        btn_storage.BackColor = Color.FromArgb(18, 109, 188)
        btn_transaction.BackColor = Color.Transparent
        btn_invoice.BackColor = Color.Transparent

        'Animasi panel
        pn_storage.Visible = True
        pn_transaction.Visible = False
        pn_invoice.Visible = False

        'Ambil data dari function
        Tampil_Data()
    End Sub

    Sub Tampil_Data()
        'Atur zoom crytal report view
        report_viewer.Zoom(1)

        'Membuat Place Holder Text
        SendMessage(Me.txt_cari_barang.Handle, &H1501, 0, "Masukan Kata Kunci")
        SendMessage(Me.txt_cari_data_transaski.Handle, &H1501, 0, "Masukan Kata Kunci")

        'Membuat ID Data
        txt_id_barang.Text = "ID.Barang-" & Format(Now, "yyMMdd.HHmmss")
        txt_no_transaksi.Text = "ID.Transaksi-" & Format(Now, "yyMMdd.HHmmss")
        txt_invoice_id.Text = "ID.Invoice-" & Format(Now, "yyMMdd.HHmmss")

        'Tampil Data Barang
        Try
            Call OpenConnection()

            Adapter = New MySqlDataAdapter("SELECT * FROM tbl_barang", db)
            DataSet = New DataSet

            Adapter.Fill(DataSet, "tbl_barang")
            dg_Barang.DataSource = DataSet.Tables("tbl_barang")
            dg_Barang.RowHeadersVisible = False
            dg_Barang.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dg_Barang.Columns(0).HeaderText = "Kode Barang"
            dg_Barang.Columns(0).Width = 150
            dg_Barang.Columns(1).HeaderText = "Nama Barang"
            dg_Barang.Columns(1).Width = 150
            dg_Barang.Columns(2).HeaderText = "Jumlah Barang"
            dg_Barang.Columns(2).Width = 150
            dg_Barang.Columns(3).HeaderText = "Harga Barang"
            dg_Barang.Columns(3).Width = 150

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()

        'Tampil Data Transaksi
        Try
            Call OpenConnection()

            Adapter = New MySqlDataAdapter("SELECT * FROM tbl_transaksi", db)
            DataSet = New DataSet

            Adapter.Fill(DataSet, "tbl_transaksi")
            dg_Transaksi.DataSource = DataSet.Tables("tbl_transaksi")
            dg_Transaksi.RowHeadersVisible = False
            dg_Transaksi.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dg_Transaksi.Columns(0).HeaderText = "No Transaksi"
            dg_Transaksi.Columns(0).Width = 150
            dg_Transaksi.Columns(1).HeaderText = "Nama Transaksi"
            dg_Transaksi.Columns(1).Width = 150
            dg_Transaksi.Columns(2).HeaderText = "Pemasukan"
            dg_Transaksi.Columns(2).Width = 150
            dg_Transaksi.Columns(3).HeaderText = "Pengeluaran"
            dg_Transaksi.Columns(3).Width = 150
            dg_Transaksi.Columns(4).HeaderText = "Waktu Transaksi"
            dg_Transaksi.Columns(4).Width = 150

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()

    End Sub

    Private Sub pn_menu_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pn_menu.MouseMove
        'Animasi form
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.Location += Control.MousePosition - Pos
        End If
        Pos = Control.MousePosition
    End Sub

    Private Sub pb_logo_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pb_logo.MouseMove
        'Animasi form
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.Location += Control.MousePosition - Pos
        End If
        Pos = Control.MousePosition
    End Sub

    Private Sub lbl_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_exit.Click
        Hapus_Data_Faktur()
        Application.Exit()
    End Sub

    Private Sub frm_main_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Hapus_Data_Faktur()
        Application.Exit()
    End Sub

    Private Sub lbl_minimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_minimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub lbl_exit_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_exit.MouseEnter
        lbl_exit.ForeColor = Color.DarkRed
    End Sub

    Private Sub lbl_exit_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_exit.MouseLeave
        lbl_exit.ForeColor = Color.White
    End Sub

    Private Sub lbl_minimize_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_minimize.MouseEnter
        lbl_minimize.ForeColor = Color.DarkCyan
    End Sub

    Private Sub lbl_minimize_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_minimize.MouseLeave
        lbl_minimize.ForeColor = Color.White
    End Sub

    Private Sub lbl_logo_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'Animasi form
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.Location += Control.MousePosition - Pos
        End If
        Pos = Control.MousePosition
    End Sub

    Private Sub btn_storage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_storage.Click
        'Animasi button menu
        btn_storage.BackColor = Color.FromArgb(18, 109, 188)
        btn_transaction.BackColor = Color.Transparent
        btn_invoice.BackColor = Color.Transparent

        'Animasi panel
        pn_storage.Visible = True
        pn_transaction.Visible = False
        pn_invoice.Visible = False

        'Ambil data dari function
        Tampil_Data()
    End Sub

    Private Sub btn_transaction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_transaction.Click
        'Animasi button menu
        btn_storage.BackColor = Color.Transparent
        btn_transaction.BackColor = Color.FromArgb(18, 109, 188)
        btn_invoice.BackColor = Color.Transparent

        'Animasi panel
        pn_storage.Visible = False
        pn_transaction.Visible = True
        pn_invoice.Visible = False

        'Ambil data dari function
        Tampil_Data()
    End Sub

    Private Sub btn_invoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_invoice.Click
        'Animasi button menu
        btn_storage.BackColor = Color.Transparent
        btn_transaction.BackColor = Color.Transparent
        btn_invoice.BackColor = Color.FromArgb(18, 109, 188)

        'Animasi panel
        pn_storage.Visible = False
        pn_transaction.Visible = False
        pn_invoice.Visible = True

        'Ambil data dari function
        Tampil_Data()
    End Sub

    Sub Simpan_Data_Barang()
        Call OpenConnection()
        Dim sql As String
        Try
            sql = "Insert into tbl_barang values (@id_barang,@nama_barang,@jumlah_barang,@harga_barang)"
            cmd = New MySqlCommand(sql, db)
            cmd.Parameters.Add(New MySqlParameter("@id_barang", txt_id_barang.Text))
            cmd.Parameters.Add(New MySqlParameter("@nama_barang", txt_nama_barang.Text))
            cmd.Parameters.Add(New MySqlParameter("@jumlah_barang", txt_jumlah_barang.Text))
            cmd.Parameters.Add(New MySqlParameter("@harga_barang", txt_harga_barang.Text))
            cmd.ExecuteNonQuery()
            MsgBox("Data Telah Tersimpan")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()
    End Sub

    Sub Clear_Data_Barang()
        txt_id_barang.Text = ""
        txt_nama_barang.Text = ""
        txt_jumlah_barang.Text = ""
        txt_harga_barang.Text = ""
        txt_cari_barang.Text = ""
    End Sub

    Sub Cari_Data_Barang()
        Try
            Call OpenConnection()
            cmd = New MySqlCommand("SELECT * FROM tbl_barang where " & _
                                   "id_barang  = '" & txt_id_barang.Text & "'", db)
            rs = cmd.ExecuteReader
            If rs.HasRows = True Then
                rs.Read()
                txt_nama_barang.Text = rs("nama_barang")
                txt_jumlah_barang.Text = rs("jumlah_barang")
                txt_harga_barang.Text = rs("harga_barang")
            Else
                MsgBox("DATA TIDAK DITEMUKAN !")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()
    End Sub

    Sub Edit_Data_Barang()
        Try
            Call OpenConnection()
            Dim sql As String
            sql = "Update tbl_barang " & _
                "Set nama_barang=@nama_barang, jumlah_barang=@jumlah_barang, harga_barang=@harga_barang " & _
                "where id_barang ='" & txt_id_barang.Text & "'"
            cmd = New MySqlCommand(sql, db)
            cmd.Parameters.Add(New MySqlParameter("@nama_barang", txt_nama_barang.Text))
            cmd.Parameters.Add(New MySqlParameter("@jumlah_barang", txt_jumlah_barang.Text))
            cmd.Parameters.Add(New MySqlParameter("@harga_barang", txt_harga_barang.Text))
            cmd.ExecuteNonQuery()
            MsgBox("Data Telah Teredit")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()
    End Sub

    Sub Hapus_Data_Barang()
        Try
            Call OpenConnection()
            Dim str As String
            str = "Delete from tbl_barang" & _
               " where id_barang = '" & txt_id_barang.Text & "'"

            cmd = New MySqlCommand(str, db)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Penghapusan Data Berhasil Dilakukan")
        Catch ec As Exception
            MessageBox.Show("Error: " & ec.Message)
        End Try
        db.Close()
    End Sub

    Private Sub btn_simpan_barang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_simpan_barang.Click
        If (String.IsNullOrEmpty(txt_id_barang.Text)) Or (String.IsNullOrEmpty(txt_nama_barang.Text)) _
            Or (String.IsNullOrEmpty(txt_jumlah_barang.Text)) Or (String.IsNullOrEmpty(txt_harga_barang.Text)) Then
            MsgBox("Data masih kosong !")
        Else
            Simpan_Data_Barang()
            Clear_Data_Barang()
            Tampil_Data()
        End If
    End Sub

    Private Sub btn_cari_barang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cari_barang.Click
        If (String.IsNullOrEmpty(txt_id_barang.Text)) Then
            MsgBox("Kode Barang masih kosong !")
        Else
            Cari_Data_Barang()
        End If
    End Sub

    Private Sub btn_edit_barang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_edit_barang.Click
        If (String.IsNullOrEmpty(txt_id_barang.Text)) Or (String.IsNullOrEmpty(txt_nama_barang.Text)) _
            Or (String.IsNullOrEmpty(txt_jumlah_barang.Text)) Or (String.IsNullOrEmpty(txt_harga_barang.Text)) Then
            MsgBox("Data masih kosong !")
        Else
            Edit_Data_Barang()
            Clear_Data_Barang()
            Tampil_Data()
        End If
    End Sub

    Private Sub btn_hapus_barang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_hapus_barang.Click
        If (String.IsNullOrEmpty(txt_id_barang.Text)) Then
            MsgBox("Data masih kosong !")
        Else
            Hapus_Data_Barang()
            Clear_Data_Barang()
            Tampil_Data()
        End If
    End Sub

    Private Sub btn_sefresh_barang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_sefresh_barang.Click
        Clear_Data_Barang()
        Tampil_Data()
    End Sub

    Private Sub btn_cari_data_barang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cari_data_barang.Click
        'Cari Data
        Try
            Call OpenConnection()

            Adapter = New MySqlDataAdapter("SELECT * FROM tbl_barang where " & _
                "id_barang like '%" & txt_cari_barang.Text.Replace("'", "''") & _
                "%' or nama_barang like '%" & txt_cari_barang.Text.Replace("'", "''") & "%'", db)
            DataSet = New DataSet

            Adapter.Fill(DataSet, "tbl_barang")
            dg_Barang.DataSource = DataSet.Tables("tbl_barang")
            dg_Barang.RowHeadersVisible = False
            dg_Barang.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dg_Barang.Columns(0).HeaderText = "Kode Barang"
            dg_Barang.Columns(0).Width = 150
            dg_Barang.Columns(1).HeaderText = "Nama Barang"
            dg_Barang.Columns(1).Width = 150
            dg_Barang.Columns(2).HeaderText = "Jumlah Barang"
            dg_Barang.Columns(2).Width = 150
            dg_Barang.Columns(3).HeaderText = "Harga Barang"
            dg_Barang.Columns(3).Width = 150
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()
    End Sub

    Private Sub btn_extract_barang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_extract_barang.Click
        If dg_Barang.RowCount > 0 Then
            'Deklarasi Object
            Dim ApExcel As Object

            'Set sebagai excel  object
            ApExcel = CreateObject("Excel.application")

            'Menyembunyikan proses excel
            ApExcel.Visible = False

            'Membuat/menambah workbook baru
            ApExcel.Workbooks.Add()

            'Lebar Kolom
            ApExcel.Columns(2).ColumnWidth = 20
            ApExcel.Columns(3).ColumnWidth = 20
            ApExcel.columns(4).columnwidth = 20
            ApExcel.columns(5).columnwidth = 20
            ApExcel.columns(6).columnwidth = 20

            'Tulis nama kolom ke excel
            For i As Integer = 2 To dg_Barang.Columns.Count + 1
                ApExcel.Cells(3, i).Value = dg_Barang.Columns(i - 2).HeaderText
            Next

            Dim strCurrency = Chr(34) & "Rp. " & vbTab & Chr(34)

            'Tulis data ke excel
            For r = 0 To dg_Barang.RowCount - 1
                For i As Integer = 2 To dg_Barang.Columns.Count + 1
                    ApExcel.Cells(r + 4, i).Value = dg_Barang.Rows(r).Cells(i - 2).Value
                Next
            Next

            Const xlCenter = -4108

            'Membuat Font Bold
            ApExcel.Range("B3:E3").Font.Bold = True

            'Membuat Font Alignmnet
            ApExcel.Range("B3:E3").VerticalAlignment = xlCenter
            ApExcel.Range("B3:E3").HorizontalAlignment = xlCenter

            'Memberi warna backgound
            ApExcel.Range("B3:E3").interior.colorindex = 51

            'Memberi warna teks
            ApExcel.Range("B3:E3").Font.colorindex = 2

            'Agar nilai cell yang panjang menjadi beberapa baris
            ApExcel.Range("B3:E" & dg_Barang.RowCount + 2).WrapText = True

            'Membuat border hitam
            ApExcel.Range("B3:E" & dg_Barang.RowCount + 3).Borders.Color = RGB(0, 0, 0)

            'Membuat Tulisan Laporan
            ApExcel.Cells(1, 4).Value = "LAPORAN DATA BARANG"

            'Membuat center aligmnet Tulisan Laporan
            ApExcel.Cells(1, 4).HorizontalAlignment = xlCenter

            'Membuat bold Tulisan Laporan
            ApExcel.Cells(1, 4).Font.Bold = True

            'Mengatur Font Size Tulisan Laporan
            ApExcel.Cells(1, 4).Font.Size = 18

            'Format Currency Excel
            ApExcel.Range("E4:E" & dg_Barang.RowCount + 3).NumberFormat = String.Format("{0}#,##0.00_);({0}#,##0.00)", strCurrency)

            ApExcel.Visible = True

            ApExcel = Nothing
        End If
    End Sub

    Sub Simpan_Data_Transaksi(ByVal no_transaksi As String, ByVal nama_transaksi As String, _
                               ByVal pemasukan_transaksi As String, ByVal pengeluaran_transaksi As String, _
                               ByVal waktu_transaksi As String)
        Call OpenConnection()
        Dim sql As String
        Try
            sql = "Insert into tbl_transaksi values (@id_transaksi,@nama_transaksi,@pemasukan_transaksi,@pengeluaran_transaksi,@waktu_transaksi)"
            cmd = New MySqlCommand(sql, db)
            cmd.Parameters.Add(New MySqlParameter("@id_transaksi", no_transaksi))
            cmd.Parameters.Add(New MySqlParameter("@nama_transaksi", nama_transaksi))
            cmd.Parameters.Add(New MySqlParameter("@pemasukan_transaksi", pemasukan_transaksi))
            cmd.Parameters.Add(New MySqlParameter("@pengeluaran_transaksi", pengeluaran_transaksi))
            cmd.Parameters.Add(New MySqlParameter("@waktu_transaksi", waktu_transaksi))
            cmd.ExecuteNonQuery()
            MsgBox("Data Telah Tersimpan")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()
    End Sub

    Sub Clear_Data_Transaksi()
        txt_no_transaksi.Text = ""
        txt_nama_transaksi.Text = ""
        txt_pemasukan_transaksi.Text = ""
        txt_pengeluaran_transaksi.Text = ""
        txt_cari_data_transaski.Text = ""
        cb_transaksi_keyword.Checked = False
        cb_transaksi_waktu.Checked = False
    End Sub

    Sub Cari_Data_Transaksi()
        Try
            Call OpenConnection()
            cmd = New MySqlCommand("SELECT * FROM tbl_transaksi where " & _
                                   "id_transaksi = '" & txt_no_transaksi.Text & "'", db)
            rs = cmd.ExecuteReader
            If rs.HasRows = True Then
                rs.Read()
                txt_nama_transaksi.Text = rs("nama_transaksi")
                txt_pemasukan_transaksi.Text = rs("pemasukan_transaksi")
                txt_pengeluaran_transaksi.Text = rs("pengeluaran_transaksi")
                dt_waktu_transaksi.Value = rs("waktu_transaksi")
            Else
                MsgBox("DATA TIDAK DITEMUKAN !")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()
    End Sub

    Sub Edit_Data_Transaksi()
        Try
            Call OpenConnection()
            Dim sql As String
            sql = "Update tbl_transaksi " & _
                "Set nama_transaksi=@nama_transaksi, pemasukan_transaksi=@pemasukan_transaksi, pengeluaran_transaksi=@pengeluaran_transaksi ,waktu_transaksi=@waktu_transaksi " & _
                "where id_transaksi ='" & txt_no_transaksi.Text & "'"
            cmd = New MySqlCommand(sql, db)
            cmd.Parameters.Add(New MySqlParameter("@nama_transaksi", txt_nama_transaksi.Text))
            cmd.Parameters.Add(New MySqlParameter("@pemasukan_transaksi", txt_pemasukan_transaksi.Text))
            cmd.Parameters.Add(New MySqlParameter("@pengeluaran_transaksi", txt_pengeluaran_transaksi.Text))
            cmd.Parameters.Add(New MySqlParameter("@waktu_transaksi", dt_waktu_transaksi.Value.ToString("yyyy-MM-dd")))
            cmd.ExecuteNonQuery()
            MsgBox("Data Telah Teredit")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()
    End Sub

    Sub Hapus_Data_Transaksi()
        Try
            Call OpenConnection()
            Dim str As String
            str = "Delete from tbl_transaksi" & _
               " where id_transaksi = '" & txt_no_transaksi.Text & "'"

            cmd = New MySqlCommand(str, db)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Penghapusan Data Berhasil Dilakukan")
        Catch ec As Exception
            MessageBox.Show("Error: " & ec.Message)
        End Try
        db.Close()
    End Sub

    Private Sub btn_simpan_transaksi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_simpan_transaksi.Click
        If (String.IsNullOrEmpty(txt_no_transaksi.Text)) Or (String.IsNullOrEmpty(txt_nama_transaksi.Text)) _
            Or (String.IsNullOrEmpty(txt_pemasukan_transaksi.Text)) Or (String.IsNullOrEmpty(txt_pengeluaran_transaksi.Text)) Then
            MsgBox("Data masih kosong !")
        Else
            Simpan_Data_Transaksi(txt_no_transaksi.Text, txt_nama_transaksi.Text, txt_pemasukan_transaksi.Text, txt_pengeluaran_transaksi.Text, dt_waktu_transaksi.Value.ToString("yyyy-MM-dd"))
            Clear_Data_Transaksi()
            Tampil_Data()
        End If
    End Sub

    Private Sub btn_cari_transaksi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cari_transaksi.Click
        If (String.IsNullOrEmpty(txt_no_transaksi.Text)) Then
            MsgBox("Kode Barang masih kosong !")
        Else
            Cari_Data_Transaksi()
        End If
    End Sub

    Private Sub btn_edit_transaksi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_edit_transaksi.Click
        If (String.IsNullOrEmpty(txt_no_transaksi.Text)) Or (String.IsNullOrEmpty(txt_nama_transaksi.Text)) _
            Or (String.IsNullOrEmpty(txt_pemasukan_transaksi.Text)) Or (String.IsNullOrEmpty(txt_pengeluaran_transaksi.Text)) Then
            MsgBox("Data masih kosong !")
        Else
            Edit_Data_Transaksi()
            Clear_Data_Transaksi()
            Tampil_Data()
        End If
    End Sub

    Private Sub btn_hapus_transaksi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_hapus_transaksi.Click
        If (String.IsNullOrEmpty(txt_no_transaksi.Text)) Then
            MsgBox("Data masih kosong !")
        Else
            Hapus_Data_Transaksi()
            Clear_Data_Transaksi()
            Tampil_Data()
        End If
    End Sub

    Private Sub btn_refresh_transaksi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_refresh_transaksi.Click
        Clear_Data_Transaksi()
        Tampil_Data()
    End Sub

    Private Sub btn_cari_data_transaksi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cari_data_transaksi.Click
        If cb_transaksi_keyword.Checked = True Then
            'Cari Data by Keyword
            Try
                Call OpenConnection()

                Adapter = New MySqlDataAdapter("SELECT * FROM tbl_transaksi where " & _
                    "id_transaksi like '%" & txt_cari_data_transaski.Text.Replace("'", "''") & _
                    "%' or nama_transaksi like '%" & txt_cari_data_transaski.Text.Replace("'", "''") & "%'", db)
                DataSet = New DataSet

                Adapter.Fill(DataSet, "tbl_transaksi")
                dg_Transaksi.DataSource = DataSet.Tables("tbl_transaksi")
                dg_Transaksi.RowHeadersVisible = False
                dg_Transaksi.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                dg_Transaksi.Columns(0).HeaderText = "No Transaksi"
                dg_Transaksi.Columns(0).Width = 150
                dg_Transaksi.Columns(1).HeaderText = "Nama Transaksi"
                dg_Transaksi.Columns(1).Width = 150
                dg_Transaksi.Columns(2).HeaderText = "Pemasukan"
                dg_Transaksi.Columns(2).Width = 150
                dg_Transaksi.Columns(3).HeaderText = "Pengeluaran"
                dg_Transaksi.Columns(3).Width = 150
                dg_Transaksi.Columns(4).HeaderText = "Waktu Transaksi"
                dg_Transaksi.Columns(4).Width = 150
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            db.Close()
        ElseIf cb_transaksi_waktu.Checked = True Then
            'Cari Data by Waktu
            Try
                Call OpenConnection()

                Dim waktu_awal As String = dt_waktu_transaksi_awal.Value.ToString("yyyy-MM-dd")
                Dim waktu_akhir As String = dt_waktu_transaksi_akhir.Value.ToString("yyyy-MM-dd")

                Adapter = New MySqlDataAdapter("SELECT * FROM tbl_transaksi where " & _
                    "waktu_transaksi between '" & waktu_awal & "' and '" & waktu_akhir & "'", db)
                DataSet = New DataSet

                Adapter.Fill(DataSet, "tbl_transaksi")
                dg_Transaksi.DataSource = DataSet.Tables("tbl_transaksi")
                dg_Transaksi.RowHeadersVisible = False
                dg_Transaksi.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                dg_Transaksi.Columns(0).HeaderText = "No Transaksi"
                dg_Transaksi.Columns(0).Width = 150
                dg_Transaksi.Columns(1).HeaderText = "Nama Transaksi"
                dg_Transaksi.Columns(1).Width = 150
                dg_Transaksi.Columns(2).HeaderText = "Pemasukan"
                dg_Transaksi.Columns(2).Width = 150
                dg_Transaksi.Columns(3).HeaderText = "Pengeluaran"
                dg_Transaksi.Columns(3).Width = 150
                dg_Transaksi.Columns(4).HeaderText = "Waktu Transaksi"
                dg_Transaksi.Columns(4).Width = 150
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            db.Close()
        ElseIf cb_transaksi_keyword.Checked = True And cb_transaksi_waktu.Checked = True Then
            'Cari Data by keyword dan Waktu
            Try
                Call OpenConnection()

                Dim waktu_awal As String = dt_waktu_transaksi_awal.Value.ToString("yyyy-MM-dd")
                Dim waktu_akhir As String = dt_waktu_transaksi_akhir.Value.ToString("yyyy-MM-dd")

                Adapter = New MySqlDataAdapter("SELECT * FROM tbl_transaksi where " & _
                    "nama_transaksi = '" & txt_cari_data_transaski.Text & "' and " & _
                    "waktu_transaksi between '" & waktu_awal & "' and '" & waktu_akhir & "'", db)
                DataSet = New DataSet

                Adapter.Fill(DataSet, "tbl_transaksi")
                dg_Transaksi.DataSource = DataSet.Tables("tbl_transaksi")
                dg_Transaksi.RowHeadersVisible = False
                dg_Transaksi.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                dg_Transaksi.Columns(0).HeaderText = "No Transaksi"
                dg_Transaksi.Columns(0).Width = 150
                dg_Transaksi.Columns(1).HeaderText = "Nama Transaksi"
                dg_Transaksi.Columns(1).Width = 150
                dg_Transaksi.Columns(2).HeaderText = "Pemasukan"
                dg_Transaksi.Columns(2).Width = 150
                dg_Transaksi.Columns(3).HeaderText = "Pengeluaran"
                dg_Transaksi.Columns(3).Width = 150
                dg_Transaksi.Columns(4).HeaderText = "Waktu Transaksi"
                dg_Transaksi.Columns(4).Width = 150
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            db.Close()
        Else
            MsgBox("Jenis pencarian belum dipilih !")
        End If


    End Sub

    Private Sub txt_excel_transaksi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_excel_transaksi.Click
        If dg_Transaksi.RowCount > 0 Then
            'Deklarasi Object
            Dim ApExcel As Object

            'Set sebagai excel  object
            ApExcel = CreateObject("Excel.application")

            'Menyembunyikan proses excel
            ApExcel.Visible = False

            'Membuat/menambah workbook baru
            ApExcel.Workbooks.Add()

            'Lebar Kolom
            ApExcel.Columns(2).ColumnWidth = 20
            ApExcel.Columns(3).ColumnWidth = 20
            ApExcel.columns(4).columnwidth = 20
            ApExcel.columns(5).columnwidth = 20
            ApExcel.columns(6).columnwidth = 20

            'Tulis nama kolom ke excel
            For i As Integer = 2 To dg_Transaksi.Columns.Count + 1
                ApExcel.Cells(3, i).Value = dg_Transaksi.Columns(i - 2).HeaderText
            Next

            Dim strCurrency = Chr(34) & "Rp. " & vbTab & Chr(34)

            'Tulis data ke excel
            For r = 0 To dg_Transaksi.RowCount - 1
                For i As Integer = 2 To dg_Transaksi.Columns.Count + 1
                    ApExcel.Cells(r + 4, i).Value = dg_Transaksi.Rows(r).Cells(i - 2).Value
                Next
            Next

            Const xlCenter = -4108

            'Membuat Font Bold
            ApExcel.Range("B3:F3").Font.Bold = True

            'Membuat Font Alignmnet
            ApExcel.Range("B3:F3").VerticalAlignment = xlCenter
            ApExcel.Range("B3:F3").HorizontalAlignment = xlCenter

            'Memberi warna backgound
            ApExcel.Range("B3:F3").interior.colorindex = 51

            'Memberi warna teks
            ApExcel.Range("B3:F3").Font.colorindex = 2

            'Agar nilai cell yang panjang menjadi beberapa baris
            ApExcel.Range("B3:F" & dg_Transaksi.RowCount + 2).WrapText = True

            'Membuat border hitam
            ApExcel.Range("B3:F" & dg_Transaksi.RowCount + 3).Borders.Color = RGB(0, 0, 0)

            'Membuat Tulisan Laporan
            ApExcel.Cells(1, 4).Value = "LAPORAN DATA TRANSASKSI"

            'Membuat center aligmnet Tulisan Laporan
            ApExcel.Cells(1, 4).HorizontalAlignment = xlCenter

            'Membuat bold Tulisan Laporan
            ApExcel.Cells(1, 4).Font.Bold = True

            'Mengatur Font Size Tulisan Laporan
            ApExcel.Cells(1, 4).Font.Size = 18

            'Format Currency Excel
            ApExcel.Range("D4:E" & dg_Transaksi.RowCount + 3).NumberFormat = String.Format("{0}#,##0.00_);({0}#,##0.00)", strCurrency)

            ApExcel.Visible = True

            ApExcel = Nothing
        End If
    End Sub

    Sub Refresh_CrystalReport_View()
        Me.Cursor = Cursors.WaitCursor

        report_viewer.ReuseParameterValuesOnRefresh = True ' Do not ask for new parameters

        report_viewer.Refresh()
        report_viewer.RefreshReport()

        Me.Cursor = Cursors.Default
    End Sub

    Sub Simpan_Data_Faktur()
        Call Connection()
        Dim sql As String
        Try
            sql = "Insert into tbl_invoice values (@id_invoice,@customer_name,@job_descriotion,@unit_price,@line_total)"
            COMMANDD = New OleDbCommand(sql, CONN)
            COMMANDD.Parameters.Add(New OleDbParameter("@id_invoice", txt_invoice_id.Text))
            COMMANDD.Parameters.Add(New OleDbParameter("@customer_name", txt_invoice_nama_pelanggan.Text))
            COMMANDD.Parameters.Add(New OleDbParameter("@job_descriotion", txt_invoice_job_description.Text))
            COMMANDD.Parameters.Add(New OleDbParameter("@unit_price", txt_invoice_unit_price.Text))
            COMMANDD.Parameters.Add(New OleDbParameter("@line_total", txt_invoice_line_total.Text))
            COMMANDD.ExecuteNonQuery()
            MsgBox("Data Telah Tersimpan")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        CONN.Close()
    End Sub

    Sub Hapus_Data_Faktur()
        Try
            Call Connection()
            Dim str As String
            str = "Delete from tbl_invoice"

            COMMANDD = New OleDbCommand(str, CONN)
            COMMANDD.ExecuteNonQuery()
            'MessageBox.Show("Penghapusan Data Berhasil Dilakukan")
        Catch ec As Exception
            MessageBox.Show("Error: " & ec.Message)
        End Try
        CONN.Close()
    End Sub

    Sub Clear_Invoice()
        txt_invoice_id.Text = ""
        txt_invoice_job_description.Text = ""
        txt_invoice_line_total.Text = ""
        txt_invoice_nama_pelanggan.Text = ""
        txt_invoice_unit_price.Text = ""
    End Sub

    Dim id_invoice, nama_invoice, waktu_invoice As String
    Dim pemasukan_invoice As Long

    Sub Simpan_Temp_Invoice()
        id_invoice = txt_invoice_id.Text
        nama_invoice = nama_invoice & "- " & txt_invoice_job_description.Text & vbCrLf
        pemasukan_invoice += Val(txt_invoice_line_total.Text)
        waktu_invoice = Format(Now, "yyyy-MM-dd")


    End Sub

    Private Sub btn_invoice_tambah_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_invoice_tambah.Click
        If (String.IsNullOrEmpty(txt_invoice_id.Text)) Or (String.IsNullOrEmpty(txt_invoice_job_description.Text)) _
            Or (String.IsNullOrEmpty(txt_invoice_line_total.Text)) Or (String.IsNullOrEmpty(txt_invoice_nama_pelanggan.Text)) _
            Or (String.IsNullOrEmpty(txt_invoice_unit_price.Text)) Then
            MsgBox("Data masih kosong !")
        Else
            Simpan_Data_Faktur()
            Simpan_Temp_Invoice()
            txt_invoice_job_description.Text = ""
            txt_invoice_line_total.Text = ""
            txt_invoice_unit_price.Text = ""
            Tampil_Data()
            Refresh_CrystalReport_View()
        End If
    End Sub

    Private Sub btn_invoice_bersih_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_invoice_bersih.Click
        Hapus_Data_Faktur()
        Clear_Invoice()
        Tampil_Data()
        Refresh_CrystalReport_View()
    End Sub

    Private Sub btn_invoice_simpan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_invoice_simpan.Click
        If Not (String.IsNullOrEmpty(id_invoice)) Or (String.IsNullOrEmpty(nama_invoice)) _
            Or (String.IsNullOrEmpty(pemasukan_invoice)) Or (String.IsNullOrEmpty(waktu_invoice)) Then
            Simpan_Data_Transaksi(id_invoice, nama_invoice, _
                                  pemasukan_invoice, "0", _
                                  waktu_invoice)
            Hapus_Data_Faktur()
            Clear_Invoice()
            Tampil_Data()
            Refresh_CrystalReport_View()
        Else
            MsgBox("Data masih kosong !")
        End If
    End Sub
End Class