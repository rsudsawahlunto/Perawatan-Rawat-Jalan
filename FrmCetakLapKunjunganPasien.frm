VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmCetakLapKunjunganPasien 
   Caption         =   "Cetak Lap Kunjungan Pasien"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "FrmCetakLapKunjunganPasien.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   5865
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
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
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmCetakLapKunjunganPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New CrDaftarKunjMskBjnsBst
Dim rptwilayah As New CrDaftarKunjBerdasarkanWilayah
Dim rptkonplngthn As New CrDaftarKunjPasienBKonPlngPTahun
Dim rptkonplng As New CrDaftarKunjBKonPlngStatus
Dim RptKonPlngTotal As New CrDaftarKunjBKonPlngStatusPerTotal
Dim rpttahunwilayah As New CrLaporanKunjunganPasienPertahun
Dim rpttahun As New CrLaporanKunjunganPasienPertahuntoWilayah
Dim Report1 As New CrDaftarKunjunganPasienBDiagnosa
Dim RptDiag As New CrDaftarKunjunganPasienBDiagnosaPerTahun
Dim report2 As New CrDaftarKunjMskBjnsBstTahun
Dim RptTotal As New cr_KunjunganPasien
Dim RptWilTot As New CrLaporanKunjunganPasienPerTotal
Dim report4 As New CrDaftarKunjunganPerDokterTh
Dim report3 As New CrDaftarKunjunganPerDokter
Dim report5 As New CrDaftarKunjunganPerDokterTotal

Dim Judul1 As String
Dim Judul2 As String
Dim Judul3 As String
Dim judul4 As String
Dim Judul5 As String
Dim Judul6 As String
Dim Judul7 As String
Dim judul8 As String
Dim judul9 As String
Dim judul10 As String
Dim judul11 As String
Dim judul12 As String
Dim Judul13 As String
Dim judul14 As String
Dim judul15 As String
Dim Judul16 As String
Dim judul17 As String
Dim judul18 As String
Dim judul19 As String

Private Sub Form_Load()
    Judul1 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN "
    Judul2 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN"
    Judul3 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN "
    judul4 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN "
    Judul5 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT PASIEN "
    Judul6 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT PASIEN"
    Judul7 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI "
    judul8 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI"
    judul9 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS PELAYANAN"
    judul10 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KELAS PELAYANAN"
    judul11 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG "
    judul12 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG "
    Judul13 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS PASIEN "
    judul14 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN RUJUKAN "
    judul15 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KASUS PENYAKIT "
    Judul16 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN JENIS OPERASI"
    judul17 = "LAPORAN KUNJUNGAN PASIEN BERDASARKAN STATUS DAN KONDISI PULANG "
    judul18 = "REKAPITULASI PASIEN PER DOKTER "
    judul19 = "REKAPITULASI PASIEN KONSUL PER DOKTER "

'    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Select Case strCetak2
        Case "LapKunjunganJenisStatusHari"
            Call KunjunganBjenisBStatusHari
        Case "LapKunjunganJenisStatusBulan"
            Call KunjunganBjenisBStatusBulan
        Case "LapKunjunganJenisStatusTahun"
            Call KunjunganBjenisBStatusTahun
        Case "LapKunjunganJenisStatusTotal"
            Call KunjunganBJenisBStatusTotal
            '======================================
        Case "LapKunjunganSt_PnyktPsnHari"
            Call LapKunjunganSt_PnyktPsnHari
        Case "LapKunjunganSt_PnyktPsnBulan"
            Call LapKunjunganSt_PnyktPsnBulan
        Case "LapKunjunganSt_PnyktPsnTahun"
            Call LapKunjunganSt_PnyktPsnTahun
        Case "LapKunjunganSt_PnyktPsnTotal"
            Call LapKunjunganSt_PnyktPsnTotal
            '==========================================
        Case "LapKunjunganBwilayahHari"
            Call LapKunjunganBwilayahHari
        Case "LapKunjunganBwilayahBulan"
            Call LapKunjunganBwilayahBulan
        Case "LapKunjunganBwilayahTahun"
            Call LapKunjunganBwilayahTahun
        Case "LapKunjunganBwilayahTotal"
            Call LapKunjunganBwilayahTotal
            '=======================================
        Case "LapKunjunganKelasStatushari"
            Call LapKunjunganKelasStatushari
        Case "LapKunjunganKelasStatusBulan"
            Call LapKunjunganKelasStatusBulan
        Case "LapKunjunganKelasStatusTahun"
            Call LapKunjunganKelasStatusTahun
        Case "LapKunjunganKelasStatusTotal"
            Call LapKunjunganKelasStatusTotal
            '=======================================
        Case "LapKunjunganRujukanBStatusHari"
            Call LapKunjunganRujukanBStatusHari
        Case "LapKunjunganRujukanBStatusBulan"
            Call LapKunjunganRujukanBStatusBulan
        Case "LapKunjunganRujukanBStatusTahun"
            Call LapKunjunganRujukanBStatusTahun
        Case "LapKunjunganRujukanBStatusTotal"
            Call LapKunjunganRujukanBStatusTotal
            '=======================================
        Case "LapKunjunganKonPulang_StatusHari"
            Call LapKunjunganKonPulang_StatusHari
        Case "LapKunjunganKonPulang_StatusBulan"
            Call LapKunjunganKonPulang_StatusBulan
        Case "LapKunjunganKonPulang_StatusTahun"
            Call LapKunjunganKonPulang_StatusTahun
        Case "LapKunjunganKonPulang_StatusTotal"
            Call LapKunjunganKonPulang_StatusTotal
            '=======================================
        Case "LapKunjunganJenisOperasi_StatusHari"
            Call LapKunjunganJenisOperasi_StatusHari
        Case "LapKunjunganJenisOperasi_StatusBulan"
            Call LapKunjunganJenisOperasi_StatusBulan
        Case "LapKunjunganJenisOperasi_StatusTahun"
            Call LapKunjunganJenisOperasi_StatusTahun
            '================================================
        Case "LapKunjunganBjenisTindakanHari"
            Call LapKunjunganBjenisTindakanHari
            '++++++++++++++++++++++++++++++++++++++++++++++++
        Case "LapKunjunganBDiagnosaHari"
            Call LapKunjunganBDiagnosaHari
        Case "LapKunjunganBDiagnosaBulan"
            Call LapKunjunganBDiagnosaBulan
        Case "LapKunjunganBDiagnosaTahun"
            Call LapKunjunganBDiagnosaTahun
        Case "LapKunjunganBDiagnosaTotal"
            Call LapKunjunganBDiagnosaTotal
            '================================================
        Case "LapKunjunganPasienBDiagnosaWilayahHari"
            Call LapKunjunganPasienBDiagnosaWilayahHari
        Case "LapKunjunganPasienBDiagnosaWilayahBulan"
            Call LapKunjunganPasienBDiagnosaWilayahBulan
        Case "LapKunjunganPasienBDiagnosaWilayahTahun"
            Call LapKunjunganPasienBDiagnosaWilayahTahun
        Case "LapKunjunganPasienBDiagnosaWilayahTotal"
            Call LapKunjunganPasienBDiagnosaWilayahTotal
            '==================================================
        Case "LapKunjunganTriaseStatusHari"
            Call LapKunjunganTriaseStatusHari
        Case "LapKunjunganTriaseStatusBulan"
            Call LapKunjunganTriaseStatusBulan
        Case "LapKunjunganTriaseStatusTahun"
            Call LapKunjunganTriaseStatusTahun
        Case "LapKunjunganTriaseStatusTotal"
            Call LapKunjunganTriaseStatusTotal
            '==================================================
        Case "LapKunjunganPerDokterHari"
            Call LapKunjunganPerDokterHari
        Case "LapKunjunganPerDokterBulan"
            Call LapKunjunganPerDokterBulan
        Case "LapKunjunganPerDokterTahun"
            Call LapKunjunganPerDokterTahun
        Case "LapKunjunganPerDokterTotal"
            Call LapKunjunganPerDokterTotal
            '==================================================
        Case "LapKunjunganKonsulPerDokterHari"
            Call LapKunjunganKonsulPerDokterHari
        Case "LapKunjunganKonsulPerDokterBulan"
            Call LapKunjunganKonsulPerDokterBulan
        Case "LapKunjunganKonsulPerDokterTahun"
            Call LapKunjunganKonsulPerDokterTahun
        Case "LapKunjunganKonsulPerDokterTotal"
            Call LapKunjunganKonsulPerDokterTotal
    End Select
End Sub

'LAPORAN PASIEN PER DOKTER
Private Sub LapKunjunganPerDokterBulan()
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report3
        .Database.AddADOCommand dbConn, adocomd
        If strCetak2 = "LapKunjunganPerDokterBulan" Then
            If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
            End If
        Else
            If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
            End If
        End If
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UnTGL.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.Dokter}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul18
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
    End With
    Dim adoKeterangan As New ADODB.Command
    Set adoKeterangan = Nothing
    Me.WindowState = 2
    adoKeterangan.ActiveConnection = dbConn
    adoKeterangan.CommandText = "select JenisPasien,Singkatan from KelompokPasien"
    adoKeterangan.CommandType = adCmdText

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report3
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report3.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'LAPORAN PASIEN KONSUL PER DOKTER
Private Sub LapKunjunganKonsulPerDokterBulan()
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report3
        .Database.AddADOCommand dbConn, adocomd
        If strCetak2 = "LapKunjunganKonsulPerDokterBulan" Then
            If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
            End If
        Else
            If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
            End If
        End If
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UnTGL.SetUnboundFieldSource ("{ado.TglDirujuk}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul19
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
    End With
    Dim adoKeterangan As New ADODB.Command
    Set adoKeterangan = Nothing
    Me.WindowState = 2
    adoKeterangan.ActiveConnection = dbConn
    adoKeterangan.CommandText = "select JenisPasien,Singkatan from KelompokPasien"
    adoKeterangan.CommandType = adCmdText

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report3
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report3.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganPerDokterHari()
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With report3
        .Database.AddADOCommand dbConn, adocomd
        If strCetak2 = "LapKunjunganPerDokterHari" Then
            If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
            End If
        Else
            If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
            End If
        End If
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UnTGL.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.Dokter}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul18
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
    End With
    Dim adoKeterangan As New ADODB.Command
    Set adoKeterangan = Nothing
    Me.WindowState = 2
    adoKeterangan.ActiveConnection = dbConn
    adoKeterangan.CommandText = "select JenisPasien,Singkatan from KelompokPasien"
    adoKeterangan.CommandType = adCmdText

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report3
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report3.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'KONSUL
Private Sub LapKunjunganKonsulPerDokterHari()
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report3
        .Database.AddADOCommand dbConn, adocomd
        If strCetak2 = "LapKunjunganKonsulPerDokterHari" Then
            If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
            End If
        Else
            If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
            End If
        End If
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UnTGL.SetUnboundFieldSource ("{ado.TglDirujuk}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.JK}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul19
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
    End With
    Dim adoKeterangan As New ADODB.Command
    Set adoKeterangan = Nothing
    Me.WindowState = 2
    adoKeterangan.ActiveConnection = dbConn
    adoKeterangan.CommandText = "select JenisPasien,Singkatan from KelompokPasien"
    adoKeterangan.CommandType = adCmdText

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report3
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report3.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganPerDokterTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report4
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.Dokter}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")

        .txtjudul.SetText judul18
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport report4, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Dim adoKeterangan As New ADODB.Command
    Set adoKeterangan = Nothing
    Me.WindowState = 2
    adoKeterangan.ActiveConnection = dbConn
    adoKeterangan.CommandText = "select JenisPasien,Singkatan from KelompokPasien"
    adoKeterangan.CommandType = adCmdText

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report4
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report4.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'KONSUL
Private Sub LapKunjunganKonsulPerDokterTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report4
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglDirujuk}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul19
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport report4, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Dim adoKeterangan As New ADODB.Command
    Set adoKeterangan = Nothing
    Me.WindowState = 2
    adoKeterangan.ActiveConnection = dbConn
    adoKeterangan.CommandText = "select JenisPasien,Singkatan from KelompokPasien"
    adoKeterangan.CommandType = adCmdText

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report4
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report4.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'TOTAL PERDOKTER
Private Sub LapKunjunganPerDokterTotal()
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report5
        .Database.AddADOCommand dbConn, adocomd
        If strCetak2 = "LapKunjunganPerDokterHari" Then
            If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "dd MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
            End If
        Else
            If CStr(Format(mdTglAwal, "dd MMMM yyyy")) = CStr(Format(mdTglAkhir, "dd MMMM yyyy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "dd MMMM yyyy")))
            End If
        End If
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .URuangan.SetUnboundFieldSource ("{ado.Dokter}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul18
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
    End With
    Dim adoKeterangan As New ADODB.Command
    Set adoKeterangan = Nothing
    Me.WindowState = 2
    adoKeterangan.ActiveConnection = dbConn
    adoKeterangan.CommandText = "select JenisPasien,Singkatan from KelompokPasien"
    adoKeterangan.CommandType = adCmdText

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report5
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report5.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'TOTAL PASIEN KONSUL PERDOKTER
Private Sub LapKunjunganKonsulPerDokterTotal()
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report5
        .Database.AddADOCommand dbConn, adocomd
        If strCetak2 = "LapKunjunganPerDokterHari" Then
            If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "dd MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
            End If
        Else
            If CStr(Format(mdTglAwal, "dd MMMM yyyy")) = CStr(Format(mdTglAkhir, "dd MMMM yyyy")) Then
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
            Else
                .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "dd MMMM yyyy")))
            End If
        End If
        .txtinstalasi.SetText mstrNamaRuangan
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .URuangan.SetUnboundFieldSource ("{ado.DokterPemeriksa}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul18
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
    End With
    Dim adoKeterangan As New ADODB.Command
    Set adoKeterangan = Nothing
    Me.WindowState = 2
    adoKeterangan.ActiveConnection = dbConn
    adoKeterangan.CommandText = "select JenisPasien,Singkatan from KelompokPasien"
    adoKeterangan.CommandType = adCmdText

    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report5
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report5.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganTriaseStatusTotal()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .txttotal.SetText "Jumlah Total Pasien " & .Priode.Text
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN KUNJUNGAN KUNJUNGAN PASIEN BERDASARKAN TRIASE DAN STATUS PASIEN ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptTotal
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptTotal.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganTriaseStatusHari()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN KUNJUNGAN KUNJUNGAN PASIEN BERDASARKAN TRIASE DAN STATUS PASIEN ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganTriaseStatusBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail

        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN TRIASE DAN STATUS PASIEN")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganTriaseStatusTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rpttahun
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN TRIASE DAN STATUS PASIEN")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rpttahun
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rpttahun.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKelasStatusTotal()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txttotal.SetText "Jumlah Total Pasien " & .Priode.Text
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN KELAS DAN STATUS PASIEN ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptTotal
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptTotal.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKelasStatushari()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN KELAS DAN STATUS PASIEN ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKelasStatusBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN KELAS DAN STATUS PASIEN ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKelasStatusTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rpttahun
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN KELAS DAN STATUS PASIEN")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rpttahun
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rpttahun.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganBwilayahTotal()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptWilTot
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN WILAYAH ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport RptWilTot, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas

    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptWilTot
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptWilTot.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganBwilayahHari()
    Call openConnection

    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rptwilayah
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN WILAYAH ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport rptwilayah, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rptwilayah
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rptwilayah.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganBwilayahBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rptwilayah
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")

        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN WILAYAH ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport rptwilayah, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rptwilayah
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rptwilayah.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganBwilayahTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rpttahunwilayah
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")

        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN WILAYAH ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport rpttahunwilayah, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rpttahunwilayah
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rpttahunwilayah.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganPasienBDiagnosaWilayahTotal()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptDiag
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.tglperiksa}")
        .UsKddiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.NamaKecamatan}")
        .UsKasus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .txttema.SetText ("Wilayah/Kecamatan")
        .UJml.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN WILAYAH DIAGNOSA")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport RptDiag, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptDiag
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptDiag.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganPasienBDiagnosaWilayahHari()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report1
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.tglperiksa}")
        .UsKddiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.NamaKecamatan}")
        .UsKasus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .txttema.SetText ("Wilayah/Kecamatan")
        .UJml.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN WILAYAH DIAGNOSA")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport Report1, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report1
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report1.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganPasienBDiagnosaWilayahBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report1
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.tglperiksa}")
        .UsKddiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.NamaKecamatan}")
        .UsKasus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .usJK.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .UJml.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txttema.SetText ("Wilayah/Kecamatan")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN WILAYAH DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport Report1, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report1
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report1.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganPasienBDiagnosaWilayahTahun()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With report2
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .Udate.SetUnboundFieldSource ("{ado.tglperiksa}")
        .UsKddiagnosa.SetUnboundFieldSource ("{ado.KdDiagnosa}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .usDiagnosa.SetUnboundFieldSource ("{ado.NamaKecamatan}")
        .txttema.SetText ("Wilayah/Kecamatan")
        .UsStatusKasus.SetUnboundFieldSource ("{ado.StatusKasus}")
        .Ujk.SetUnboundFieldSource ("{ado.JenisKelamin}")
        .UJml.SetUnboundFieldSource ("{ado.JmlKunjungan}")
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN WILAYAH DIAGNOSA")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport report2, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report2
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report2.PrintOut False
        Unload Me
    End If
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
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
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
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport report2, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = report2
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        report2.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganBDiagnosaBulan()
'    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report1
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
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
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport Report1, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report1
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report1.PrintOut False
        Unload Me
    End If
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
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
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
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport RptDiag, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptDiag
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptDiag.PrintOut False
        Unload Me
    End If
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
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
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
        .txtjudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS DIAGNOSA ")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "1"
            sDuplex = "0"
        End If
        settingreport Report1, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report1
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report1.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjunganBJenisBStatusTotal()
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txttotal.SetText "Jumlah Total Pasien " & .Priode.Text
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.JK}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptTotal
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptTotal.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjunganBjenisBStatusHari()
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail

        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.JK}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul1
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjunganBjenisBStatusBulan()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.JK}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul2
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub KunjunganBjenisBStatusTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rpttahun
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.JK}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul13
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rpttahun
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rpttahun.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganSt_PnyktPsnTotal()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txttotal.SetText "Jumlah Total Pasien " & .Priode.Text
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul5
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptTotal
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptTotal.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'=================================== Berdasarkan kunjungan Kasus Penyakit & Status Pasien ================
Private Sub LapKunjunganSt_PnyktPsnHari()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul5
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganSt_PnyktPsnBulan()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul6
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganSt_PnyktPsnTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rpttahun
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul15
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rpttahun
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rpttahun.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'===================================Kunjungan Berdasarkan Rujukan Dan setatus ====================
Private Sub LapKunjunganRujukanBStatusTotal()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptTotal
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txttotal.SetText "Jumlah Total Pasien " & .Priode.Text
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul3
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport RptTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptTotal
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptTotal.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganRujukanBStatusHari()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul3
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganRujukanBStatusBulan()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul4
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganRujukanBStatusTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rpttahun
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul14
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rpttahun
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rpttahun.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'===================================Kunjungan Berdasarkan Status & Kondisi Pulang ====================
Private Sub LapKunjunganKonPulang_StatusTotal()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With RptKonPlngTotal
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtKet1.SetText ("SM       :Sembuh")
        .txtket2.SetText ("BJ       :Berobat Jalan")
        .txtket3.SetText ("C        :Cacat")
        .txtket4.SetText ("M<48     :Meninggal < 48 Jam")
        .txtket5.SetText ("M>48     :Meninggal > 48 Jam")
        .txtket6.SetText ("L        :Lain - Lain")
        .txtket7.SetText ("RI       :Dirawat Inap")
        .txtket8.SetText ("Ref      :Referal")
        .txtket9.SetText ("DOA      :Death Of Arrived")
        .txtket10.SetText ("M IGD     :Meninggal di IGD")

        .txttotal.SetText "Jumlah Total Pasien " & .Priode.Text
        .Udate.SetUnboundFieldSource ("{ado.tglkeluar}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul11
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport RptKonPlngTotal, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = RptKonPlngTotal
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        RptKonPlngTotal.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKonPulang_StatusHari()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rptkonplng
        .Database.AddADOCommand dbConn, adocomd
        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtKet1.SetText ("SM       :Sembuh")
        .txtket2.SetText ("BJ       :Berobat Jalan")
        .txtket3.SetText ("C        :Cacat")
        .txtket4.SetText ("M<48     :Meninggal < 48 Jam")
        .txtket5.SetText ("M>48     :Meninggal > 48 Jam")
        .txtket6.SetText ("L        :Lain - Lain")
        .txtket7.SetText ("RI       :Dirawat Inap")
        .txtket8.SetText ("Ref      :Referal")
        .txtket9.SetText ("DOA      :Death Of Arrived")
        .txtket10.SetText ("M IGD     :Meninggal di IGD")

        .Udate.SetUnboundFieldSource ("{ado.tglkeluar}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul11
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport rptkonplng, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rptkonplng
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rptkonplng.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKonPulang_StatusBulan()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rptkonplng
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If

        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtKet1.SetText ("SM       :Sembuh")
        .txtket2.SetText ("BJ       :Berobat Jalan")
        .txtket3.SetText ("C        :Cacat")
        .txtket4.SetText ("M<48     :Meninggal < 48 Jam")
        .txtket5.SetText ("M>48     :Meninggal > 48 Jam")
        .txtket6.SetText ("L        :Lain - Lain")
        .txtket7.SetText ("RI       :Dirawat Inap")
        .txtket8.SetText ("Ref      :Referal")
        .txtket9.SetText ("DOA      :Death Of Arrived")
        .txtket10.SetText ("M IGD     :Meninggal di IGD")

        .Udate.SetUnboundFieldSource ("{ado.tglkeluar}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul12
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport rptkonplng, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rptkonplng
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rptkonplng.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganKonPulang_StatusTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rptkonplngthn
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtKet1.SetText ("SM       :Sembuh")
        .txtket2.SetText ("BJ       :Berobat Jalan")
        .txtket3.SetText ("C        :Cacat")
        .txtket4.SetText ("M<48     :Meninggal < 48 Jam")
        .txtket5.SetText ("M>48     :Meninggal > 48 Jam")
        .txtket6.SetText ("L        :Lain - Lain")
        .txtket7.SetText ("RI       :Dirawat Inap")
        .txtket8.SetText ("Ref      :Referal")
        .txtket9.SetText ("DOA      :Death Of Arrived")
        .txtket10.SetText ("M IGD     :Meninggal di IGD")
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.tglkeluar}")
        .usJudul.SetUnboundFieldSource ("{ado.Judul}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsDetail.SetUnboundFieldSource ("{ado.Detail}")
        .usJK.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul17
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport rptkonplngthn, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rptkonplngthn
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rptkonplngthn.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'===================================Kunjungan Berdasarkan Status & Jenis Oprasi ====================
Private Sub LapKunjunganJenisOperasi_StatusHari()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul7
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganJenisOperasi_StatusBulan()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "MMMM yyyy")) = CStr(Format(mdTglAkhir, "MMMM yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText judul8
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganJenisOperasi_StatusTahun()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With rpttahun
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "yyyy")) = CStr(Format(mdTglAkhir, "yyyy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul16
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = rpttahun
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        rpttahun.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

'================================================== Jenis Tindakan =======================
Private Sub LapKunjunganBjenisTindakanHari()
    Call openConnection
    Set FrmCetakLapKunjunganPasien = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL

    With Report
        .Database.AddADOCommand dbConn, adocomd

        If CStr(Format(mdTglAwal, "mm-dd-yy")) = CStr(Format(mdTglAkhir, "mm-dd-yy")) Then
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd MMMM yyyy")))
        Else
            .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "dd-MM-yy")) & " s/d " & CStr(Format(mdTglAkhir, "dd-MM-yy")))
        End If
        .txtNamaRS.SetText strNNamaRS
        .txtalamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .Udate.SetUnboundFieldSource ("{ado.TglPendaftaran}")
        .UJudul.SetUnboundFieldSource ("{ado.Judul}")
        .URuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UDetail.SetUnboundFieldSource ("{ado.Detail}")
        .Ujk.SetUnboundFieldSource ("{ado.Jk}")
        .UJml.SetUnboundFieldSource ("{ado.JmlPasien}")
        .txtjudul.SetText Judul2
    End With

    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ReportSource = Report
            .EnableGroupTree = False
            .ViewReport
            .Zoom 1
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmCetakLapKunjunganPasien = Nothing
    sUkuranKertas = ""
End Sub
