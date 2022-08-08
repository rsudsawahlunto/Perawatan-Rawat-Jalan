VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmViewerLaporan 
   Caption         =   "Viewer Laporan"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   Icon            =   "FrmViewerLaporan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   10845
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
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
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "FrmViewerLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reportTopten As New crDiagnosaTopTen2
Dim reporrtoptengrafik As New crDiagnosaTopTenGrafik
Dim reportBukuBesar As New crBukuBesar
Dim rptRekapKunjunganPIGrafik As New crRekapKunjunganPIGrafik
Dim repLapPasienKonsul As New crLapPasienKonsul

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    adocomd.ActiveConnection = dbConn

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Dim tanggal As String

    Select Case cetak

            'Rekapitulasi 10 besar Penyakit
        Case "RekapTopten"
            adocomd.CommandText = "sELECT * FROM V_RekapitulasiDiagnosaTopTen " _
            & "WHERE (TglPeriksa BETWEEN '" _
            & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') " _
            & " and NamaRuangan = '" & mstrNamaRuangan & "' ORDER BY instalasi,diagnosa"

            adocomd.CommandType = adCmdText
            reportTopten.Database.AddADOCommand dbConn, adocomd

            If Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy") Then
                tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
            Else
                tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
            End If

            With reportTopten
                .Text1.SetText strNNamaRS
                .Text2.SetText strNAlamatRS
                .Text3.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .txtPeriode2.SetText tanggal
                .txtinstalasi.SetText ""
                .usSMF.SetUnboundFieldSource ("{ado.instalasi}")
                .usDiagnosa.SetUnboundFieldSource ("{ado.diagnosa}")
                .unJumlahPasien.SetUnboundFieldSource ("{ado.jumlahpasien}")
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport reportTopten, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
            End With

            CRViewer1.ReportSource = reportTopten

        Case "RekapToptenGrafik"
            adocomd.CommandText = "sELECT * FROM V_RekapitulasiDiagnosaTopTen " _
            & "WHERE (TglPeriksa BETWEEN '" _
            & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "')  " _
            & " and NamaRuangan = '" & mstrNamaRuangan & "' ORDER BY instalasi,diagnosa"

            adocomd.CommandType = adCmdText
            reporrtoptengrafik.Database.AddADOCommand dbConn, adocomd

            If Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy") Then
                tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
            Else
                tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
            End If

            With reporrtoptengrafik
                .Text1.SetText strNNamaRS
                .Text2.SetText strNAlamatRS
                .Text3.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .txtPeriode2.SetText tanggal
                .txtinstalasi.SetText ""
                .usSMF.SetUnboundFieldSource ("{ado.instalasi}")
                .usDiagnosa.SetUnboundFieldSource ("{ado.diagnosa}")
                .unJumlahPasien.SetUnboundFieldSource ("{ado.jumlahpasien}")
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport reporrtoptengrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
            End With

            CRViewer1.ReportSource = reporrtoptengrafik

            'Data rekapitulasi kasus penyakit

        Case "BukuBesar"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With reportBukuBesar
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtTgl.SetText Format(FrmBukuRegister.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmBukuRegister.DTPickerAkhir, "dd/MM/yyyy")
                .udtTglMasuk.SetUnboundFieldSource "{ado.tglmasuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.NoRegister}"
                .usCM.SetUnboundFieldSource "{ado.nocm}"
                .usPasien.SetUnboundFieldSource "{ado.namapasien}"
                .usAlamat.SetUnboundFieldSource "{ado.alamat}"
                .usUmur.SetUnboundFieldSource "{ado.umur}"
                .usJK.SetUnboundFieldSource "{ado.jk}"
                .usRujukan.SetUnboundFieldSource "{ado.AsalRujukan}"
                .usDiagnosa.SetUnboundFieldSource "{ado.diagnosa}"
                .usKlpkPasien.SetUnboundFieldSource "{ado.jenispasien}"
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport reportBukuBesar, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = reportBukuBesar

        Case "RekapKunjunganPIGrafik"
            adocomd.CommandText = "sELECT * FROM V_RekapitulasiPasienBRujukanInternal " _
            & "WHERE (TglPendaftaran BETWEEN '" _
            & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" _
            & " AND kdruangan = '" & mstrKdRuangan & "'"

            adocomd.CommandType = adCmdText
            rptRekapKunjunganPIGrafik.Database.AddADOCommand dbConn, adocomd

            If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
                tanggal = "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy")
            Else
                tanggal = "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
            End If

            With rptRekapKunjunganPIGrafik
                .Text1.SetText strNNamaRS
                .Text2.SetText strNAlamatRS
                .Text3.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .TxtTanggal.SetText tanggal
                .txtinstalasi.SetText ""
                .Text5.SetText strNamaRuangan
                .uskdruangan.SetUnboundFieldSource ("{ado.kdruangan}")
                .usRuangan.SetUnboundFieldSource ("{ado.ruangan}")
                .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
                .unjmllaki.SetUnboundFieldSource ("{ado.jmlpasienpria}")
                .unjmlperempuan.SetUnboundFieldSource ("{ado.jmlpasienwanita}")
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport rptRekapKunjunganPIGrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
            End With
            CRViewer1.ReportSource = rptRekapKunjunganPIGrafik

            'PASIEN KONSUL
        Case "PasienKonsul"
            adocomd.CommandText = strSQL
            adocomd.CommandType = adCmdText
            With repLapPasienKonsul
                .Text16.SetText strNNamaRS
                .Text18.SetText strNAlamatRS
                .Text19.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .Database.AddADOCommand dbConn, adocomd
                .txtRuang.SetText strNNamaRuangan
                .txtinstalasi.SetText xDaftarInstalasiA
                .txtTgl.SetText Format(FrmBukuRegister.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmBukuRegister.DTPickerAkhir, "dd/MM/yyyy")

                .udTgl.SetUnboundFieldSource "{ado.TglDirujuk}"
                .usNoDaf.SetUnboundFieldSource "{ado.nopendaftaran}"
                .usNoCM.SetUnboundFieldSource "{ado.NoCM}"
                .usNmPasien.SetUnboundFieldSource "{ado.NamaPasien}"
                .usAlamat.SetUnboundFieldSource "{ado.Alamat}"
                .usUmur.SetUnboundFieldSource "{ado.Umur}"
                .usJK.SetUnboundFieldSource "{ado.JK}"
                .usAslRujukan.SetUnboundFieldSource "{ado.RuangPerujuk}"
                .usDiagnosa.SetUnboundFieldSource "{ado.Diagnosa}"
                .usKelompokPsn.SetUnboundFieldSource "{ado.JenisPasien}"
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport repLapPasienKonsul, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = repLapPasienKonsul
    End Select
    Screen.MousePointer = vbHourglass
    If vLaporan = "view" Then
        With CRViewer1
            .ViewReport
            .Zoom 1
        End With
    Else
        If cetak = "RekapTopten" Then
            reportTopten.PrintOut False
            Unload Me
        ElseIf cetak = "RekapToptenGrafik" Then
            reporrtoptengrafik.PrintOut False
            Unload Me
        ElseIf cetak = "BukuBesar" Then
            reportBukuBesar.PrintOut False
            Unload Me
        ElseIf cetak = "RekapKunjunganPIGrafik" Then
            rptRekapKunjunganPIGrafik.PrintOut False
            Unload Me
        ElseIf cetak = "PasienKonsul" Then
            repLapPasienKonsul.PrintOut False
            Unload Me
        End If
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
    Set FrmViewerLaporan = Nothing
End Sub

