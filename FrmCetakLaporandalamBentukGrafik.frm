VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakLaporandalamBentukGrafik 
   Caption         =   "Cetak Laporan dalam Bentuk Grafik"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   Icon            =   "FrmCetakLaporandalamBentukGrafik.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   9555
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmCetakLaporandalamBentukGrafik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New Cr_RekapGrafik
Dim ReportPerTahun As New Cr_RekapGrafikPerTahun
Dim RptTotal As New Cr_RekapGrafikPerTotal
Dim Judul1 As String
Dim Report1 As New CrDaftarKunjunganPasienBDiagnosa
Dim report2 As New CrDaftarKunjMskBjnsBstTahun
Dim RptDiag As New CrDaftarKunjunganPasienBDiagnosaPerTahun

Private Sub Form_Load()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Select Case strCetak2
        Case "LapKunjunganJenisStatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganJenisStatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganJenisStatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganJenisStatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT  "
            Call RekapKunjunganPerTotal

            '======================================
        Case "LapKunjunganSt_PnyktPsnHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganSt_PnyktPsnBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganSt_PnyktPsnTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT  "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganSt_PnyktPsnTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT  "
            Call RekapKunjunganPerTotal

            '==========================================
        Case "LapKunjunganBwilayahHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN WILAYAH  "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganBwilayahBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS WILAYAH  "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganBwilayahTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS WILAYAH "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganBwilayahTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS WILAYAH "
            Call RekapKunjunganPerTotal

            '=======================================
        Case "LapKunjunganKelasStatushari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS  "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganKelasStatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS  "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganKelasStatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS "
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganKelasStatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS  "
            Call RekapKunjunganPerTotal

            '=======================================
        Case "LapKunjunganRujukanBStatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN "
            Call KunjunganBjenisBStatusHari
        Case "LapKunjunganRujukanBStatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN"
            Call KunjunganBjenisBStatusBulan
        Case "LapKunjunganRujukanBStatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN"
            Call KunjunganBjenisBStatusTahun

        Case "LapKunjunganRujukanBStatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN "
            Call RekapKunjunganPerTotal

            '=======================================
        Case "LapKunjunganKonPulang_StatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG  "
            Call LapKunjunganKonPulang_StatusHari

        Case "LapKunjunganKonPulang_StatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG  "
            Call LapKunjunganKonPulang_StatusBulan

        Case "LapKunjunganKonPulang_StatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG "
            Call LapKunjunganKonPulang_StatusTahun

        Case "LapKunjunganKonPulang_StatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG "
            
            '=======================================
        Case "LapKunjunganJenisOperasi_StatusHari"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
            Call KunjunganBjenisBStatusHari

        Case "LapKunjunganJenisOperasi_StatusBulan"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
            Call KunjunganBjenisBStatusBulan

        Case "LapKunjunganJenisOperasi_StatusTahun"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
            Call KunjunganBjenisBStatusTahun

            '================================================
        Case "LapKunjunganBjenisTindakanHari"
            Call KunjunganBjenisBStatusHari
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
        Case "LapKunjunganBDiagnosaHari"
            Judul1 = "GRAFIK REKAPITULASI PASIEN BERDASARKAN DIAGNOSA"
            Call LapPasienBDiagnosaGrafikPerHari
        Case "LapKunjunganBDiagnosaBulan"
            Judul1 = "GRAFIK REKAPITULASI PASIEN BERDASARKAN DIAGNOSA"
            Call LapPasienBDiagnosaGrafikPerBulan
        Case "LapKunjunganBDiagnosaTahun"
            Judul1 = "GRAFIK REKAPITULASI PASIEN BERDASARKAN DIAGNOSA"
            Call LapPasienBDiagnosaGrafikPerTahun
        Case "LapKunjunganBDiagnosaTotal"
            Judul1 = "GRAFIK REKAPITULASI PASIEN BERDASARKAN DIAGNOSA"
            Call LapPasienBDiagnosaGrafikPerTotal
        Case "LapKunjunganBDiagnosaHari"
            Call LapKunjunganBDiagnosaHari
        Case "LapKunjunganBDiagnosaBulan"
            Call LapKunjunganBDiagnosaBulan
        Case "LapKunjunganBDiagnosaTahun"
            Call LapKunjunganBDiagnosaTahun
        Case "LapKunjunganBDiagnosaTotal"
            Call LapKunjunganBDiagnosaTotal

            '==================================================
        Case "LapKunjunganTriaseStatusHari"
            Call KunjunganBjenisBStatusHari
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN TRIASE  "
        Case "LapKunjunganTriaseStatusBulan"
            Call KunjunganBjenisBStatusBulan
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN TRIASE  "
        Case "LapKunjunganTriaseStatusTahun"
            Call KunjunganBjenisBStatusTahun
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN TRIASE "
        Case "LapKunjunganTriaseStatusTotal"
            Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN TRIASE  "
            Call RekapKunjunganPerTotal
        Case "LapKunjunganPasienBDiagnosaWilayahHari"
            Judul1 = "GRAFIK REKAPITULASI PASIEN BERDASARKAN WILAYAH DIAGNOSA"
            LapPasienBWilayahDiagnosaGrafikPerHari
        Case "LapKunjunganPasienBDiagnosaWilayahBulan"
            Judul1 = "GRAFIK REKAPITULASI PASIEN BERDASARKAN WILAYAH DIAGNOSA"
            LapPasienBWilayahDiagnosaGrafikPerBulan
        Case "LapKunjunganPasienBDiagnosaWilayahTahun"
            Judul1 = "GRAFIK REKAPITULASI PASIEN BERDASARKAN WILAYAH DIAGNOSA"
            LapPasienBWilayahDiagnosaGrafikPerTahun
        Case "LapKunjunganPasienBDiagnosaWilayahTotal"
            Judul1 = "GRAFIK REKAPITULASI PASIEN BERDASARKAN WILAYAH DIAGNOSA"
            LapPasienBWilayahDiagnosaGrafikPerTotal
    End Select
End Sub

Private Sub RekapKunjunganPerTotal()
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        If rs.RecordCount < 100 Then
            .Graph1.Width = 9240
            .txtJudul.Width = 11280
            .Periode.Left = 6720
            .txtfootjudul.Width = 11280
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "1"
                sDuplex = "0"
            End If
        Else
            .Graph1.Width = 18720
            .Graph1.Left = 120
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "2"
                sDuplex = "0"
            End If
        End If

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = RptTotal
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub KunjunganBjenisBStatusHari()
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        If rs.RecordCount < 100 Then
            .Graph1.Width = 9240
            .txtJudul.Width = 11280
            .Periode.Left = 6720
            .txtfootjudul.Width = 11280
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "1"
                sDuplex = "0"
            End If
        Else
            .Graph1.Width = 18720
            .Graph1.Left = 120
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "2"
                sDuplex = "0"
            End If
        End If

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub KunjunganBjenisBStatusBulan()
    Call openConnection
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub KunjunganBjenisBStatusTahun()
    Call openConnection
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With ReportPerTahun

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport ReportPerTahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = ReportPerTahun
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub LapKunjunganKonPulang_StatusHari()
    Call openConnection
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.tglkeluar}")
        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub LapKunjunganKonPulang_StatusBulan()
    Call openConnection
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.tglkeluar}")
        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub LapKunjunganKonPulang_StatusTahun()
    Call openConnection
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With ReportPerTahun
        .Database.AddADOCommand dbConn, adocomd
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.tglkeluar}")
        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlPasien}")

        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport ReportPerTahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = ReportPerTahun
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCetakLaporandalamBentukGrafik = Nothing
    sUkuranKertas = ""
End Sub

Private Sub LapPasienBDiagnosaGrafikPerHari()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usJK.SetUnboundFieldSource ("{ado.Jeniskelamin}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub LapPasienBDiagnosaGrafikPerBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usJK.SetUnboundFieldSource ("{ado.Jeniskelamin}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapPasienBDiagnosaGrafikPerTahun()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With ReportPerTahun

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport ReportPerTahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = ReportPerTahun
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapPasienBDiagnosaGrafikPerTotal()
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        If rs.RecordCount < 100 Then
            .Graph1.Width = 9240
            .txtJudul.Width = 11280
            .Periode.Left = 6720
            .txtfootjudul.Width = 11280
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "1"
                sDuplex = "0"
            End If
        Else
            .Graph1.Width = 18720
            .Graph1.Left = 120
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "2"
                sDuplex = "0"
            End If
        End If

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText Judul1
        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = RptTotal
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganBDiagnosaHari()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report1

        .Database.AddADOCommand dbConn, adocomd

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.tglperiksa}")
        .UsKddiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.Diagnosa}")
        .UsKasus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .txttema.SetText ("Diagnosa")
        .UJml.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If

        settingreport Report1, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report1
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganBDiagnosaBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report1
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.tglperiksa}")
        .UsKddiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.Diagnosa}")
        .UsKasus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .UJml.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txttema.SetText ("Diagnosa")
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport Report1, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report1
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub LapKunjunganBDiagnosaTahun()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With report2
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.tglperiksa}")
        .UsKddiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.Diagnosa}")
        .UsStatusKasus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .Ujk.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .txttema.SetText ("Diagnosa")
        .UJml.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport report2, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = report2
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub LapKunjunganBDiagnosaTotal()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptDiag

        .Database.AddADOCommand dbConn, adocomd

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.tglperiksa}")
        .UsKddiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.Diagnosa}")
        .UsKasus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .txttema.SetText ("Diagnosa")
        .UJml.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If

        settingreport RptDiag, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = RptDiag
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapPasienBWilayahDiagnosaGrafikPerHari()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.NamaKecamatan}")
        .usJK.SetUnboundFieldSource ("{ado.Jeniskelamin}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub LapPasienBWilayahDiagnosaGrafikPerBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.NamaKecamatan}")
        .usJK.SetUnboundFieldSource ("{ado.Jeniskelamin}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapPasienBWilayahDiagnosaGrafikPerTahun()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With ReportPerTahun

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.NamaKecamatan}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport ReportPerTahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = ReportPerTahun
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapPasienBWilayahDiagnosaGrafikPerTotal()
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        If rs.RecordCount < 100 Then
            .Graph1.Width = 9240
            .txtJudul.Width = 11280
            .Periode.Left = 6720
            .txtfootjudul.Width = 11280
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "1"
                sDuplex = "0"
            End If
        Else
            .Graph1.Width = 18720
            .Graph1.Left = 120
            If sUkuranKertas = "" Then
                sUkuranKertas = "5"
                sOrientasKertas = "2"
                sDuplex = "0"
            End If
        End If

        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Periode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UTAnggal.SetUnboundFieldSource ("{ado.TglPeriksa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.NamaKecamatan}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .UnJmlPasien.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtJudul.SetText Judul1
        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = RptTotal
        .ViewReport
        .Zoom (98)
    End With
    Screen.MousePointer = vbDefault
End Sub
